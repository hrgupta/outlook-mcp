"""Inbox rule tools: list, create, update, delete message rules.

Wraps the Graph ``/me/mailFolders/inbox/messageRules`` endpoint. Rules are
server-side: they run whether or not any Outlook client is open, and they
apply to incoming mail in the Inbox. Each rule pairs a set of *conditions*
(what an incoming message must match) with a set of *actions* (what to do
when it matches).

The Graph SDK ships typed builders for this endpoint, so the wrapper uses
them directly rather than hand-rolling OData.
"""

from __future__ import annotations

import asyncio
from typing import Any

from outlook_mcp.config import Config
from outlook_mcp.errors import ToolInputError
from outlook_mcp.folder_resolver import resolve_folder_id
from outlook_mcp.permissions import CATEGORY_MAIL_RULES, check_permission, check_sends_mail
from outlook_mcp.validation import sanitize_output, validate_email, validate_graph_id


def _recipient(address: str) -> Any:
    """Build a Graph ``Recipient`` from an email address."""
    from msgraph.generated.models.email_address import EmailAddress
    from msgraph.generated.models.recipient import Recipient

    ea = EmailAddress()
    ea.address = address
    r = Recipient()
    r.email_address = ea
    return r


def _recipients(addresses: list[str]) -> list[Any]:
    return [_recipient(a) for a in addresses]


def _copy_actions(existing: Any) -> Any:
    """Copy every action off an existing rule's ``actions`` block.

    A PATCH that carries an ``actions`` object replaces the rule's actions
    wholesale — Graph does not merge this collection — so any action absent
    from the payload is cleared. Re-pointing a rule while preserving its
    mark-as-read / forward / stop-processing behaviour therefore means reading
    the current actions and sending them all back. Dropping one is exactly the
    silent-no-op shape this project's offline guards exist to catch.
    """
    from msgraph.generated.models.message_rule_actions import MessageRuleActions

    actions = MessageRuleActions()
    if existing is None:
        return actions
    # Copy only what the rule actually has. Assigning an attribute None marks it
    # changed in the backing store, and the request adapter then emits that null
    # under its *Python* name — the shape that made Graph answer 400 on a partial
    # contact address (see tests/test_write_payloads_reach_the_wire.py).
    for attr in (
        "assign_categories",
        "copy_to_folder",
        "delete",
        "forward_as_attachment_to",
        "forward_to",
        "mark_as_read",
        "mark_importance",
        "move_to_folder",
        "permanent_delete",
        "redirect_to",
        "stop_processing_rules",
    ):
        value = getattr(existing, attr, None)
        if value is not None:
            setattr(actions, attr, value)
    return actions


_FORWARDING_ACTIONS = ("forward_to", "redirect_to", "forward_as_attachment_to")

# One lock per rule id. `update_rule` is a read-modify-write: Graph replaces the
# `actions` object wholesale, so the payload is built from a read taken moments
# earlier. Two overlapping calls — a folder change and a mark-read change, the
# obvious pair — each read the old actions and each send a complete replacement,
# so the later request quietly restores the earlier caller's old value in the
# field it never mentioned. The lock is per process because that is the only
# sharing this server controls; see the docstring for the rest.
_rule_locks: dict[str, asyncio.Lock] = {}


def _rule_lock(rule_id: str) -> asyncio.Lock:
    lock = _rule_locks.get(rule_id)
    if lock is None:
        lock = asyncio.Lock()
        _rule_locks[rule_id] = lock
    return lock


def _forwards_offsite(actions: Any) -> bool:
    """True if this rule's actions send mail to an address outside the mailbox.

    Used on re-enable: a disabled rule can be switched back on without any
    action being supplied, which resumes whatever it does — including forwarding
    to a third party — so the send gate has to look at what the rule *is*, not
    only at what the caller passed.
    """
    if actions is None:
        return False
    return any(getattr(actions, attr, None) for attr in _FORWARDING_ACTIONS)


_CONDITION_FIELDS = (
    "body_contains",
    "body_or_subject_contains",
    "from_addresses",
    "has_attachments",
    "header_contains",
    "importance",
    "message_action_flag",
    "recipient_contains",
    "sender_contains",
    "sent_cc_addresses",
    "sent_to_addresses",
    "subject_contains",
    "within_size_range",
)

_ACTION_FIELDS = (
    "assign_categories",
    "copy_to_folder",
    "delete",
    "forward_as_attachment_to",
    "forward_to",
    "mark_as_read",
    "mark_importance",
    "move_to_folder",
    "permanent_delete",
    "redirect_to",
    "stop_processing_rules",
)


def _stops_processing(actions: Any) -> bool:
    """True if this rule halts evaluation, hiding the rules that follow it."""
    return bool(getattr(actions, "stop_processing_rules", None)) if actions is not None else False


def _masks_a_forwarding_rule(current: Any, rules: Any) -> bool:
    """True if `current` halts processing ahead of an enabled forwarding rule.

    A forwarding rule placed after a `stop_processing_rules` rule never runs.
    Changing *that* stop rule — disabling it, reordering it later, deleting it —
    switches the forwarding on without touching the forwarding rule at all, so
    the send gate has to look at the rules the stop rule is hiding, not only at
    the one named in the request.
    """
    if not _stops_processing(getattr(current, "actions", None)):
        return False
    current_sequence = getattr(current, "sequence", None) or 0
    for rule in rules or []:
        if getattr(rule, "id", None) == getattr(current, "id", None):
            continue
        if not getattr(rule, "is_enabled", False):
            continue
        if not _forwards_offsite(getattr(rule, "actions", None)):
            continue
        if (getattr(rule, "sequence", None) or 0) > current_sequence:
            return True
    return False


async def _rules_in_inbox(graph_client: Any) -> list[Any]:
    response = await graph_client.me.mail_folders.by_mail_folder_id("inbox").message_rules.get()
    return list(response.value) if response and response.value else []


def _address(recipient: Any) -> str:
    """The address out of a Graph ``Recipient``, scrubbed for display."""
    email = getattr(recipient, "email_address", None)
    return sanitize_output(getattr(email, "address", None) or "")


def _describe_conditions(conditions: Any) -> dict:
    """Serialize a rule's conditions into plain JSON, dropping what it does not set.

    A rule you cannot read the behaviour of cannot be audited: "is anything
    forwarding mail out of this mailbox?" is the question that matters and the
    name does not answer it. Empty sub-objects are omitted rather than emitted as
    ``{}`` so a move-only rule reads as one short block.
    """
    if conditions is None:
        return {}
    described: dict[str, Any] = {}
    for attr in _CONDITION_FIELDS:
        value = getattr(conditions, attr, None)
        if value is None or value == []:
            continue
        if attr == "from_addresses":
            described[attr] = [_address(a) for a in value]
        elif attr == "within_size_range":
            described[attr] = {
                "min": getattr(value, "minimum_size", None),
                "max": getattr(value, "maximum_size", None),
            }
        elif isinstance(value, list):
            described[attr] = [
                sanitize_output(str(v)) if isinstance(v, str) else str(v) for v in value
            ]
        else:
            described[attr] = value
    return described


def _describe_actions(actions: Any) -> dict:
    """Serialize a rule's actions into plain JSON, dropping what it does not set.

    Same reason as `_describe_conditions`, and more load-bearing: a forwarded
    address is the single most important fact about a rule someone did not
    write themselves.
    """
    if actions is None:
        return {}
    described: dict[str, Any] = {}
    for attr in _ACTION_FIELDS:
        value = getattr(actions, attr, None)
        if value is None or value == []:
            continue
        if attr in _FORWARDING_ACTIONS:
            described[attr] = [_address(a) for a in value]
        elif attr == "assign_categories":
            described[attr] = [sanitize_output(str(c)) for c in value]
        else:
            described[attr] = value
    return described


async def list_rules(graph_client: Any) -> dict:
    """List all inbox rules, sorted by execution order (sequence).

    Each rule carries its conditions and actions as plain JSON, not only its
    name and status: the reason to list rules is to find out what they *do* —
    an unexpected forwarding rule is the thing you are looking for. Rule names
    and every caller-visible string inside those blocks are scrubbed of control
    characters the way every other caller-visible string is.
    """
    response = await graph_client.me.mail_folders.by_mail_folder_id("inbox").message_rules.get()
    rules = list(response.value) if response and response.value else []
    rules.sort(key=lambda r: r.sequence if r.sequence is not None else 9999)
    return {
        "count": len(rules),
        "rules": [
            {
                "id": r.id,
                "display_name": sanitize_output(r.display_name or ""),
                "is_enabled": r.is_enabled,
                "sequence": r.sequence,
                "is_read_only": r.is_read_only,
                "conditions": _describe_conditions(getattr(r, "conditions", None)),
                "actions": _describe_actions(getattr(r, "actions", None)),
            }
            for r in rules
        ],
    }


async def create_rule(
    graph_client: Any,
    name: str,
    *,
    config: Config,
    from_addresses: list[str] | None = None,
    subject_contains: list[str] | None = None,
    body_contains: list[str] | None = None,
    sender_contains: list[str] | None = None,
    has_attachments: bool | None = None,
    importance: str | None = None,
    move_to_folder: str | None = None,
    copy_to_folder: str | None = None,
    mark_as_read: bool | None = None,
    delete: bool | None = None,
    forward_to: list[str] | None = None,
    assign_categories: list[str] | None = None,
    stop_processing_rules: bool | None = None,
    is_enabled: bool = True,
    sequence: int | None = None,
) -> dict:
    """Create an inbox rule from conditions and actions.

    At least one condition and one action are required. ``move_to_folder`` /
    ``copy_to_folder`` accept display names, well-known names, or Graph IDs.
    ``importance`` is "low", "normal", or "high". ``sequence`` controls
    execution order (lower runs first); omitted, the rule is appended after
    the highest existing sequence.

    A rule is durable, not a one-off: it keeps running after this session ends
    and applies to every message that arrives later. ``forward_to`` is also an
    email to somebody else, so it additionally requires the ``mail_send``
    category. Hard delete is deliberately not offered here — ``delete`` moves
    matching mail to Deleted Items, which is recoverable.
    """
    check_permission(config, CATEGORY_MAIL_RULES, "outlook_create_rule")
    if forward_to:
        # A forwarding rule emails a third party long after this session ended,
        # which is precisely what `mail_send` gates. Same ordering as the
        # calendar tools: the tool's own category first, then the send gate.
        check_sends_mail(config, "outlook_create_rule", "forward matching mail to another address")
    if not name or not name.strip():
        raise ToolInputError("Rule name is required.")

    from msgraph.generated.models.importance import Importance
    from msgraph.generated.models.message_rule import MessageRule
    from msgraph.generated.models.message_rule_actions import MessageRuleActions
    from msgraph.generated.models.message_rule_predicates import MessageRulePredicates

    # ── Conditions ─────────────────────────────────────────────
    conditions = MessageRulePredicates()
    has_condition = False
    if from_addresses:
        conditions.from_addresses = _recipients([validate_email(a) for a in from_addresses])
        has_condition = True
    if subject_contains:
        conditions.subject_contains = subject_contains
        has_condition = True
    if body_contains:
        conditions.body_contains = body_contains
        has_condition = True
    if sender_contains:
        conditions.sender_contains = sender_contains
        has_condition = True
    if has_attachments is not None:
        conditions.has_attachments = has_attachments
        has_condition = True
    if importance:
        valid = {"low", "normal", "high"}
        if importance not in valid:
            raise ToolInputError(f"importance must be one of {sorted(valid)}; got {importance}")
        conditions.importance = Importance(importance)
        has_condition = True

    # ── Actions ────────────────────────────────────────────────
    actions = MessageRuleActions()
    has_action = False
    if move_to_folder:
        actions.move_to_folder = await resolve_folder_id(graph_client, move_to_folder)
        has_action = True
    if copy_to_folder:
        actions.copy_to_folder = await resolve_folder_id(graph_client, copy_to_folder)
        has_action = True
    if mark_as_read is not None:
        actions.mark_as_read = mark_as_read
        has_action = True
    if delete is not None:
        actions.delete = delete
        has_action = True
    if forward_to:
        actions.forward_to = _recipients([validate_email(a) for a in forward_to])
        has_action = True
    if assign_categories:
        actions.assign_categories = assign_categories
        has_action = True
    if stop_processing_rules is not None:
        actions.stop_processing_rules = stop_processing_rules
        has_action = True

    if not has_condition:
        raise ToolInputError(
            "At least one condition is required: from_addresses, subject_contains, "
            "body_contains, sender_contains, has_attachments, or importance."
        )
    if not has_action:
        raise ToolInputError(
            "At least one action is required: move_to_folder, copy_to_folder, "
            "mark_as_read, delete, forward_to, assign_categories, or "
            "stop_processing_rules."
        )

    # ── Assemble ───────────────────────────────────────────────
    rule = MessageRule()
    rule.display_name = name.strip()
    rule.is_enabled = is_enabled
    rule.conditions = conditions
    rule.actions = actions

    if sequence is not None:
        if sequence < 1:
            raise ToolInputError("sequence must be a positive integer")
        rule.sequence = sequence
    else:
        # Append after the highest existing sequence (default 100 if none).
        existing = await list_rules(graph_client)
        highest = max((r["sequence"] for r in existing["rules"]), default=0)
        rule.sequence = max(highest + 1, 100)

    created = await graph_client.me.mail_folders.by_mail_folder_id("inbox").message_rules.post(rule)
    return {
        "status": "created",
        "id": created.id if created else None,
        "display_name": created.display_name if created else name,
        "sequence": created.sequence if created else rule.sequence,
    }


async def update_rule(
    graph_client: Any,
    rule_id: str,
    *,
    config: Config,
    name: str | None = None,
    is_enabled: bool | None = None,
    sequence: int | None = None,
    move_to_folder: str | None = None,
    mark_as_read: bool | None = None,
) -> dict:
    """Update an existing inbox rule (partial patch — only provided fields change).

    ``move_to_folder`` accepts display names, well-known names, or Graph IDs.
    Pass is_enabled=False to disable a rule without deleting it. The rule's
    other actions survive: Graph replaces the ``actions`` object rather than
    merging it, so the current actions are read and sent back with the
    requested field overridden.

    Re-enabling (``is_enabled=True``) reads the rule first: switching a
    forwarding or redirecting rule back on requires the ``mail_send`` category,
    because it resumes mailing a third party for as long as the rule exists.

    Ordering is gated too, because it can switch forwarding on without any
    forwarding argument being supplied: a forwarding rule behind a
    ``stop_processing_rules`` rule never runs, so moving that forwarding rule
    earlier, or moving / disabling the stop rule, starts delivering mail. Any
    sequence change to a forwarding rule, and any change that unmasks a later
    forwarding rule, requires ``mail_send``.

    Overlapping updates of the same rule are serialized inside this server, so
    two calls in flight from one session cannot undo each other. Across
    processes the last write wins: Graph's ``messageRules`` endpoint returns no
    ``@odata.etag`` (checked against the live endpoint), so there is no
    ``If-Match`` to reject a stale write with.
    """
    check_permission(config, CATEGORY_MAIL_RULES, "outlook_update_rule")
    rule_id = validate_graph_id(rule_id)

    from msgraph.generated.models.message_rule import MessageRule

    builder = graph_client.me.mail_folders.by_mail_folder_id(
        "inbox"
    ).message_rules.by_message_rule_id(rule_id)

    patch = MessageRule()
    if name is not None:
        patch.display_name = name.strip()
    if is_enabled is not None:
        patch.is_enabled = is_enabled
    if sequence is not None:
        if sequence < 1:
            raise ToolInputError("sequence must be a positive integer")
        patch.sequence = sequence

    # One read covers both reasons to look at the rule: re-pointing needs the
    # existing actions to send back, and re-enabling needs them to see whether
    # this rule emails somebody else.
    needs_current = (
        move_to_folder is not None
        or mark_as_read is not None
        or is_enabled is True
        # Reordering can switch forwarding on without touching the rule that
        # does it: moving a forwarding rule ahead of the stop rule that was
        # hiding it, or moving / disabling that stop rule back behind it.
        or sequence is not None
        or is_enabled is False
    )

    async with _rule_lock(rule_id):
        current = await builder.get() if needs_current else None
        current_actions = current.actions if current else None

        if is_enabled is True and _forwards_offsite(current_actions):
            check_sends_mail(config, "outlook_update_rule", "re-enable a rule that forwards mail")

        if sequence is not None and _forwards_offsite(current_actions):
            check_sends_mail(
                config,
                "outlook_update_rule",
                "move a forwarding rule in the run order, ahead of the rule hiding it",
            )

        if (is_enabled is False or sequence is not None) and _masks_a_forwarding_rule(
            current,
            await _rules_in_inbox(graph_client) if _stops_processing(current_actions) else None,
        ):
            check_sends_mail(
                config,
                "outlook_update_rule",
                "let a later forwarding rule run by changing the rule that stops processing",
            )

        if move_to_folder is not None or mark_as_read is not None:
            actions = _copy_actions(current.actions if current else None)
            if move_to_folder is not None:
                actions.move_to_folder = await resolve_folder_id(graph_client, move_to_folder)
            if mark_as_read is not None:
                actions.mark_as_read = mark_as_read
            patch.actions = actions

        updated = await builder.patch(patch)
    return {
        "status": "updated",
        "id": updated.id if updated else rule_id,
        "display_name": updated.display_name if updated else name,
        "is_enabled": updated.is_enabled if updated else is_enabled,
        "sequence": updated.sequence if updated else sequence,
    }


async def delete_rule(graph_client: Any, rule_id: str, *, config: Config) -> dict:
    """Delete an inbox rule by ID. Irreversible.

    Deleting a rule that halts processing can start a later forwarding rule, so
    the same ``mail_send`` check that guards creating or re-enabling one applies
    here when the rule being deleted is what was keeping the forwarding off.
    """
    check_permission(config, CATEGORY_MAIL_RULES, "outlook_delete_rule")
    rule_id = validate_graph_id(rule_id)

    builder = graph_client.me.mail_folders.by_mail_folder_id(
        "inbox"
    ).message_rules.by_message_rule_id(rule_id)

    async with _rule_lock(rule_id):
        current = await builder.get()
        if _stops_processing(getattr(current, "actions", None)) and _masks_a_forwarding_rule(
            current, await _rules_in_inbox(graph_client)
        ):
            check_sends_mail(
                config,
                "outlook_delete_rule",
                "let a later forwarding rule run by deleting the rule that stops processing",
            )
        await builder.delete()

    _rule_locks.pop(rule_id, None)
    return {"status": "deleted", "id": rule_id}

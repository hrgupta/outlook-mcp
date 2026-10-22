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

from typing import Any

from outlook_mcp.config import Config
from outlook_mcp.folder_resolver import resolve_folder_id
from outlook_mcp.permissions import CATEGORY_MAIL_RULES, check_permission
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


async def list_rules(graph_client: Any) -> dict:
    """List all inbox rules, sorted by execution order (sequence).

    Rule names are scrubbed of control characters the way every other
    caller-visible string is.
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
    permanent_delete: bool | None = None,
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
    """
    check_permission(config, CATEGORY_MAIL_RULES, "outlook_create_rule")
    if not name or not name.strip():
        raise ValueError("Rule name is required.")

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
            raise ValueError(f"importance must be one of {valid}; got {importance}")
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
    if permanent_delete is not None:
        actions.permanent_delete = permanent_delete
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
        raise ValueError(
            "At least one condition is required: from_addresses, subject_contains, "
            "body_contains, sender_contains, has_attachments, or importance."
        )
    if not has_action:
        raise ValueError(
            "At least one action is required: move_to_folder, copy_to_folder, "
            "mark_as_read, delete, permanent_delete, forward_to, assign_categories, "
            "or stop_processing_rules."
        )

    # ── Assemble ───────────────────────────────────────────────
    rule = MessageRule()
    rule.display_name = name.strip()
    rule.is_enabled = is_enabled
    rule.conditions = conditions
    rule.actions = actions

    if sequence is not None:
        if sequence < 1:
            raise ValueError("sequence must be a positive integer")
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
            raise ValueError("sequence must be a positive integer")
        patch.sequence = sequence
    if move_to_folder is not None or mark_as_read is not None:
        current = await builder.get()
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
    """Delete an inbox rule by ID. Irreversible."""
    check_permission(config, CATEGORY_MAIL_RULES, "outlook_delete_rule")
    rule_id = validate_graph_id(rule_id)
    await (
        graph_client.me.mail_folders.by_mail_folder_id("inbox")
        .message_rules.by_message_rule_id(rule_id)
        .delete()
    )
    return {"status": "deleted", "id": rule_id}

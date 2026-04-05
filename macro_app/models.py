from __future__ import annotations

from copy import deepcopy
from typing import Any

from .constants import ACTION_DEFINITIONS, ACTION_TEXT_INPUT, build_default_params, normalize_action_name

DEFAULT_DELAY = "2"
_default_delay = DEFAULT_DELAY
DEFAULT_WORKFLOW_NAME = "新工作流"
DEFAULT_NODE_ALIAS = "大节点"


def get_default_delay() -> str:
    return _default_delay


def set_default_delay(value: str) -> None:
    global _default_delay
    cleaned = str(value).strip() or DEFAULT_DELAY
    _default_delay = cleaned


def create_small_node(action_name: str = ACTION_TEXT_INPUT) -> dict[str, Any]:
    action_name = normalize_action_name(action_name)
    if action_name not in ACTION_DEFINITIONS:
        action_name = ACTION_TEXT_INPUT
    return {
        "alias": "",
        "action": action_name,
        "delay": get_default_delay(),
        "params": build_default_params(action_name),
    }


def create_big_node(index: int = 1) -> dict[str, Any]:
    return {
        "alias": f"{DEFAULT_NODE_ALIAS} {index}",
        "scope": {
            "regex": "",
            "bring_front": True,
            "multi_window_mode": "first",
        },
        "small_nodes": [create_small_node()],
    }


def create_workflow(name: str = DEFAULT_WORKFLOW_NAME) -> dict[str, Any]:
    return {"name": name, "big_nodes": [create_big_node(1)]}


def create_step(action_name: str = ACTION_TEXT_INPUT) -> dict[str, Any]:
    return create_small_node(action_name)


def create_node(index: int = 1) -> dict[str, Any]:
    return create_big_node(index)


def get_big_nodes(workflow: dict[str, Any]) -> list[dict[str, Any]]:
    nodes = workflow.setdefault("big_nodes", workflow.pop("nodes", [create_big_node(1)]))
    if not nodes:
        nodes.append(create_big_node(1))
    return nodes


def get_small_nodes(big_node: dict[str, Any]) -> list[dict[str, Any]]:
    steps = big_node.setdefault("small_nodes", big_node.pop("steps", [create_small_node()]))
    if not steps:
        steps.append(create_small_node())
    return steps


def _normalize_scope(scope: dict[str, Any]) -> dict[str, Any]:
    regex = (scope.get("regex") or "").strip()
    if not regex:
        keyword = (scope.get("keyword") or "").strip()
        exclude = (scope.get("exclude") or "").strip()
        if keyword and exclude:
            regex = rf"^(?=.*{keyword})(?!.*{exclude}).*$"
        elif keyword:
            regex = keyword
    mode = str(scope.get("multi_window_mode", "first") or "first").strip().lower()
    if mode not in {"first", "sync", "serial"}:
        mode = "first"
    return {
        "regex": regex,
        "bring_front": bool(scope.get("bring_front", True)),
        "multi_window_mode": mode,
    }


def normalize_small_node(small_node: dict[str, Any]) -> dict[str, Any]:
    action_name = normalize_action_name(str(small_node.get("action", ACTION_TEXT_INPUT)))

    params = small_node.get("params")
    params_dict: dict[str, Any] = params if isinstance(params, dict) else {}

    normalized_params = build_default_params(action_name)
    for field in normalized_params:
        value = params_dict.get(field, normalized_params[field])
        normalized_params[field] = "" if value is None else str(value)

    return {
        "alias": str(small_node.get("alias", "") or ""),
        "action": action_name,
        "delay": str(small_node.get("delay", small_node.get("delay_after_step", get_default_delay())) or get_default_delay()),
        "params": normalized_params,
    }


def normalize_big_node(big_node: dict[str, Any], index: int) -> dict[str, Any]:
    alias = str(big_node.get("alias") or big_node.get("name") or f"{DEFAULT_NODE_ALIAS} {index}")
    scope = _normalize_scope(big_node.get("scope", {}))
    small_nodes = [normalize_small_node(item) for item in get_small_nodes(big_node)]
    return {
        "alias": alias,
        "scope": scope,
        "small_nodes": small_nodes or [create_small_node()],
    }


def normalize_workflow(workflow: dict[str, Any]) -> dict[str, Any]:
    name = str(workflow.get("name") or DEFAULT_WORKFLOW_NAME)
    big_nodes = [normalize_big_node(item, index + 1) for index, item in enumerate(get_big_nodes(workflow))]
    return {
        "name": name,
        "big_nodes": big_nodes or [create_big_node(1)],
    }


def clone_payload(payload: dict[str, Any]) -> dict[str, Any]:
    return deepcopy(payload)

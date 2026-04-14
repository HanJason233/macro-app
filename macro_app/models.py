from __future__ import annotations

from copy import deepcopy
from typing import Any

from .constants import ACTION_DEFINITIONS, ACTION_TEXT_INPUT, build_default_params, normalize_action_name

DEFAULT_DELAY = "2"
_default_delay = DEFAULT_DELAY
DEFAULT_WORKFLOW_NAME = "新工作流"
DEFAULT_NODE_ALIAS = "大节点"
FLOW_NEXT = "next"
FLOW_LOOP = "loop"
FLOW_JUMP = "jump"
FLOW_CONDITIONAL_JUMP = "conditional_jump"
FLOW_STOP = "stop"
FLOW_CONDITIONS = {"last_detected", "last_not_detected", "window_text_detected", "window_image_detected"}


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
        "node_delay": "0",
        "scope": {
            "regex": "",
            "bring_front": True,
            "multi_window_mode": "first",
        },
        "flow": {
            "mode": FLOW_NEXT,
            "target": "",
            "condition": "last_detected",
            "max_loops": "1",
            "condition_target_text": "",
            "condition_text_match_mode": "contains",
            "condition_min_score": "0.5",
            "condition_image_path": "",
            "condition_match_mode": "default",
            "condition_threshold": "0.8",
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


def _normalize_flow(flow: dict[str, Any]) -> dict[str, Any]:
    mode = str(flow.get("mode", FLOW_NEXT) or FLOW_NEXT).strip().lower()
    if mode not in {FLOW_NEXT, FLOW_LOOP, FLOW_JUMP, FLOW_CONDITIONAL_JUMP, FLOW_STOP}:
        mode = FLOW_NEXT

    target = str(flow.get("target", "") or "").strip()
    condition = str(flow.get("condition", "last_detected") or "last_detected").strip().lower()
    if condition not in FLOW_CONDITIONS:
        condition = "last_detected"

    max_loops = str(flow.get("max_loops", "1") or "1").strip()
    try:
        max_loops_value = int(max_loops)
        if max_loops_value < 1:
            max_loops_value = 1
    except Exception:
        max_loops_value = 1

    return {
        "mode": mode,
        "target": target,
        "condition": condition,
        "max_loops": str(max_loops_value),
        "condition_target_text": str(flow.get("condition_target_text", "") or "").strip(),
        "condition_text_match_mode": str(flow.get("condition_text_match_mode", "contains") or "contains").strip().lower(),
        "condition_min_score": str(flow.get("condition_min_score", "0.5") or "0.5").strip(),
        "condition_image_path": str(flow.get("condition_image_path", "") or "").strip(),
        "condition_match_mode": str(flow.get("condition_match_mode", "default") or "default").strip().lower(),
        "condition_threshold": str(flow.get("condition_threshold", "0.8") or "0.8").strip(),
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
    node_delay = str(
        big_node.get(
            "node_delay",
            big_node.get("node_interval_delay", big_node.get("delay_between_nodes", "0")),
        )
        or "0"
    ).strip()
    scope = _normalize_scope(big_node.get("scope", {}))
    flow = _normalize_flow(big_node.get("flow", {}))
    small_nodes = [normalize_small_node(item) for item in get_small_nodes(big_node)]
    return {
        "alias": alias,
        "node_delay": node_delay,
        "scope": scope,
        "flow": flow,
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

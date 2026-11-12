"""职位名称分类与理论 post_rank 推导。

职位词表来源于 4post_rank.R。四组词表依次对应 rank1、rank2、rank3、rank4。
这里只做保守的标准化匹配；无法识别的职位会返回 unknown，交给人工处理。
"""

from __future__ import annotations

import re
from typing import Any

from .rank_adjuster import DISPLAY_RANK, RANK_SCORE, canonical_rank


DEPT_LEADERS = frozenset("书记|党组书记|主任|市长|主席|党委书记|理事长|政治委员|代理市长|市委书记|代理检察长|代理院长|检察长|院长|校长|社长|总编辑|局长|关长|站长|部长|台长|处长|监委书记|馆长|临时党委书记|临时局长|主任委员|会长|党工委书记|党总支书记|党总支部书记|团长|理事会主任|政委|第一政治委员|支队长|联合党委书记|直属机关党委书记|机关党委书记|第一书记|党组支部书记|党组总支部书记|组长|党组党组书记|党组直属机关党委书记|党组党委书记|第一政委|名誉主席|主委|队长|场长|大队长|代理主任|主席（会长）|局长（主任）|校（院）长|临时委员会书记|校长（院长）|所长|管委会主任|一级指挥长|署长".split("|"))
DEPT_DEPUTY_LEADERS = frozenset("常委|副书记|党组副书记|党组成员|副主任|副部长|副组长|成员|副司令员|副主席|常务委员|副市长|副校长|党委副书记|常务副主任|常务副市长|常务副主席|副团长|第一副书记|常务副书记|常务副校长|常务副部长|专职副书记|专职副主任|副局长|副台长|副主任委员|副院长|工会主席|副处长|纪检组长|总经济师|常务副局长|总工程师|专职副主席|副理事长|总会计师|副政委|副支队长|副政治委员|参谋长|副社长|副总编辑|理事会副主任|副关长|副站长|工委会主任|工委会主席|副总编|纪检组组长|经审委主委|纪检支队长|副主委|纪检监察组组长|常务副校（院）长|副校（院）长|纪监工委书记|兼职副主席|兼职副会长|专职委员|常务副会长|专职副会长|副局长（副主任）|总农艺师|总畜牧兽医师|林业总工程师|总规划师|副大队长|副会长|工会主任|总审计师|总统计师|专职副主任委员|临时党委副书记|副队长|总兽医师|监察员|党支部书记|副所长|纪工委书记|管委会副主任".split("|"))
SECOND_LEVEL_LEADERS = frozenset("秘书长|机关党组书记|党委委员|市委书记助理|委员".split("|"))
SECOND_LEVEL_DEPUTY_LEADERS = frozenset("常务副秘书长|副秘书长".split("|"))

POST_GROUPS = {
    "rank1": DEPT_LEADERS,
    "rank2": DEPT_DEPUTY_LEADERS,
    "rank3": SECOND_LEVEL_LEADERS,
    "rank4": SECOND_LEVEL_DEPUTY_LEADERS,
}


def normalize_post_title(value: Any) -> str:
    text = "" if value is None else str(value).strip()
    text = re.sub(r"\s+", "", text)
    return text.replace("（", "（").replace(")", "）")


def classify_post_title(value: Any) -> str | None:
    title = normalize_post_title(value)
    if not title:
        return None
    for group, titles in POST_GROUPS.items():
        if title in titles:
            return group
    return None


def expected_post_rank(wrank_t: Any, post_group: str | None) -> str | None:
    """Return the display spelling of the expected rank for a title group."""
    if post_group not in POST_GROUPS:
        return None
    base = RANK_SCORE.get(canonical_rank(wrank_t))
    if base is None:
        return None
    offset = int(post_group[-1]) - 1
    score = base - offset
    if score < 0 or score >= len(RANK_SCORE):
        return None
    canonical = next(name for name, value in RANK_SCORE.items() if value == score)
    return DISPLAY_RANK.get(canonical, canonical)


def audit_post_rank_frame(frame):
    """Return detailed pre-audit rows without changing the input frame."""
    import pandas as pd

    rows = []
    for index, row in frame.iterrows():
        title = row.get("post1", "")
        group = classify_post_title(title)
        expected = expected_post_rank(row.get("wrank_t", ""), group)
        original = "" if pd.isna(row.get("post_rank", "")) else str(row.get("post_rank", "")).strip()
        if group is None:
            status = "职位名称无法识别"
        elif expected is None:
            status = "部门级别无法判断"
        elif not original:
            status = "原 post_rank 为空"
        elif canonical_rank(original) == canonical_rank(expected):
            status = "一致"
        else:
            status = "post_rank 不一致"
        if status != "一致":
            rows.append({
                "source_index": index,
                "year": row.get("year", ""),
                "cityname": row.get("cityname", ""),
                "distname": row.get("distname", ""),
                "w_name": row.get("w_name", ""),
                "post1": title,
                "wrank_t": row.get("wrank_t", ""),
                "职位类别": group or "",
                "理论post_rank": expected or "",
                "原post_rank": original,
                "issue_type": status,
            })
    return pd.DataFrame(rows)

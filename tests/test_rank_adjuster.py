from src.engine.rank_adjuster import adjust_post_rank, canonical_rank, rank_score


def test_historical_aliases_are_canonicalized():
    assert canonical_rank("地厅级") == "正厅级"
    assert canonical_rank("厅级") == "正厅级"
    assert rank_score("处级") == rank_score("正处级")


def test_department_downgrade_moves_post_rank_down_one_level():
    result = adjust_post_rank("副厅级", "副厅级", "处级")
    assert result.adjusted == "处级"
    assert result.delta == -1
    assert result.status == "adjusted"


def test_unknown_rank_is_not_silently_changed():
    result = adjust_post_rank("初级", "副厅级", "处级")
    assert result.adjusted == "初级"
    assert result.status == "unresolved"


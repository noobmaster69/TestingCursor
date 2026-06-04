from file_convert.handlers import build_registry


def test_list_pairs_not_empty():
    registry = build_registry()
    pairs = registry.list_pairs()
    assert ("md", "html") in pairs
    assert ("xlsx", "csv") in pairs


def test_resolve_unknown_pair():
    registry = build_registry()
    try:
        registry.resolve("docx", "png")
        assert False, "expected error"
    except Exception as exc:
        assert "Cannot convert" in str(exc)

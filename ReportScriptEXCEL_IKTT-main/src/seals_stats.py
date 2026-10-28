def build_seals_stats(paths, cfg) -> dict:
    raw = load_sources(paths)
    norm = normalize_sources(raw, cfg)
    registry = build_registry(norm, cfg)
    checked = apply_rules_and_checks(registry, cfg)
    tables = build_output_tables(checked, cfg)
    return {
        "raw": raw,
        "normalized": norm,
        "registry": registry,
        "checked": checked,
        "tables": tables,
    }
def load_sources(paths):
    print("[seals_stats] load_sources: prototype")
    return []

def normalize_sources(raw, cfg):
    print("[seals_stats] normalize_sources: prototype")
    return raw

def build_registry(norm, cfg):
    print("[seals_stats] build_registry: prototype")
    return {}

def apply_rules_and_checks(registry, cfg):
    print("[seals_stats] apply_rules_and_checks: prototype")
    return registry

def build_output_tables(checked, cfg):
    print("[seals_stats] build_output_tables: prototype")
    return {}
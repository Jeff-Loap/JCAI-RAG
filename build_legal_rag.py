# -*- coding: utf-8 -*-
from legal_agent import LegalRAGStore, get_default_config


def main():
    store = LegalRAGStore(get_default_config())
    stats = store.rebuild()
    print(f"入库完成：documents={stats.documents}, chunks={stats.chunks}")
    for source in stats.sources:
        print(f"- {source}")


if __name__ == "__main__":
    main()

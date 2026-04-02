# -*- coding: utf-8 -*-
from legal_agent import LegalRAGAgent, LegalRAGStore, get_default_config


def main():
    config = get_default_config()
    store = LegalRAGStore(config)
    agent = LegalRAGAgent(store=store, config=config)
    while True:
        question = input("\n请输入问题（回车退出）：").strip()
        if not question:
            break
        result = agent.ask(question)
        print("\n回答：")
        print(result["answer"])
        print("\n引用：")
        for item in result["citations"]:
            print(f"- {item}")


if __name__ == "__main__":
    main()

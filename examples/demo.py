from pathlib import Path

from meeting_agent.app import main


examples = Path(__file__).resolve().parent
settings = ["--db", "data/demo/meetings.sqlite3", "--documents", "data/demo/documents"]
for day, title in (("2026-11-23", "客户门户上线准备"), ("2026-11-26", "客户门户灰度进展")):
    main([*settings, "ingest", str(examples / f"{day}.txt"), "--project", "portal", "--meeting", f"meeting-{day}",
          "--date", day, "--title", title, "--fixture", str(examples / f"{day}.json")])
main([*settings, "report", "--project", "portal", "--name", "客户门户项目", "--start", "2026-11-23", "--end", "2026-11-27"])

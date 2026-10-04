# Contributing

Thanks for helping improve Movie to Excel. Keep changes aligned with the project's local-first purpose: the workbook belongs to the user, personal ratings and notes stay local, and internet access is used only to verify public metadata or identify a supplied poster.

Before opening a pull request:

1. Keep `Master List` authoritative for movies. TV series and documentaries have their own archive tables.
2. Preserve all seven sheets, their filters, the grouped region layout, and user-controlled settings.
3. Do not add accounts, analytics, cloud storage, public profiles, likes, feeds, or social rankings.
4. Use established Chinese director names for Mainland China, Hong Kong, and Taiwan productions; use full English or romanized names elsewhere.
5. Run `node plugins/movie-to-excel/scripts/smoke-test.mjs`, verify the workbook, and inspect changed sheets visually.
6. Do not commit personal viewing archives, temporary previews, or generated inspection files.

Bug reports should include the title, affected sheet, expected behavior, and a screenshot with private notes removed.

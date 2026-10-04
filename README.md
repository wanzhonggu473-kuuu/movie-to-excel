# Movie to Excel

Movie to Excel is a local-first Codex plugin for keeping a private Excel archive of movies, TV series, and documentaries. Give Codex a title or poster and, after identifying the work, it can add a record to a workbook on your computer. Public metadata lookup may use the internet; your viewing history stays in the local file.

![Movie archive region overview](docs/images/movie-archive-overview.png)

## Why this project

Movie to Excel does not try to compete with Douban's ratings, reviews, social features, or community. Its purpose is to let you own a simple, searchable record that you can edit, back up, and move as an ordinary Excel file.

## 为什么做这个项目

Movie to Excel 不打算复制豆瓣的评分、评论、社交功能或用户社区。它的重点是让你在自己的电脑上保存一份可查询、可修改、可备份和迁移的观影记录。联网只用于核对公开的作品资料；个人观看历史保存在本地 Excel 文件中。

## Workbook

The included [blank template](plugins/movie-to-excel/assets/movie-archive-template.xlsx) has seven sheets:

| Sheet | Purpose |
| --- | --- |
| `My Movies` | All movies grouped across two rows of production regions. The lower row moves down when the upper groups grow. |
| `Region View` | Six horizontal region lists, each with a director filter. Europe/Oceania and Other Regions also have country filters. |
| `Master List` | The authoritative, searchable movie list, including optional private rating and notes. |
| `Statistics` | Local movie totals by region. |
| `Settings` | The workbook's display and date conventions. |
| `My TV Series` | Series title, watched seasons, viewing period, and country/region. |
| `My Documentaries` | Documentary title, director, year, seasons/episodes watched, viewing period, and country/region. |

Grouped titles show the Chinese title before the official English or original title. For Mainland China, Hong Kong, and Taiwan films, directors use their established Chinese names; for other films, directors use full English or romanized names. A documentary series may have different episode directors, so its director field can remain blank when no single verified director applies.

## New: TV series and documentaries / 新增：电视剧与纪录片

`My TV Series` records the seasons you actually watched, the viewing period (normally to the month), and the country or region. A single season is written as `S1`; more seasons can be added later to the same entry. Titles use Chinese first, followed by the official English or original title. You can simply say that a title is a TV series and provide its name, seasons, and viewing period.

`My Documentaries` is a separate list for documentaries and documentary series. It can record a director when one is verified, release year, watched seasons or episodes, viewing period, and country or region. Watching three episodes is recorded as progress such as `E1–E3`, not as completion of the whole series. If a volume or season is uncertain, the record says so rather than guessing. Later viewing updates the existing entry instead of creating another row.

`My TV Series` 用来记录实际看过的季数、观看时期（通常精确到月份）和国家或地区。只看过一季就写 `S1`，以后继续观看可以更新同一条记录。片名先写中文名，再写官方英文名或原名。告诉 Codex“这是电视剧”，再提供名称、季数和观看时期即可。

`My Documentaries` 单独保存纪录片和纪录剧集，可记录核实过的导演、年份、实际看过的季或集、观看时期及国家或地区。比如只看了前三集，就记作 `E1–E3`，不会误写成整季看完；不确定属于哪一卷或哪一季时，会保留“不确定”。以后继续看时更新原条目，不重复新增一行。

## Recording rules / 记录规则

- If the user does not explicitly say “TV series” or “documentary,” treat the title as a movie. / 没有明确说“电视剧”或“纪录片”时，默认按电影处理。
- Append a newly watched title to the end of its movie region or series list. For an existing series, update its progress in the same row. / 新条目追加到对应分类末尾；继续观看同一剧集时更新原行的进度。
- Keep the viewing date or period at the precision the user provides: `26-09-28`, `26-09`, `26`, `2609`, or a range. Blank means unknown. When no movie date is supplied at all, use today's local date and state that assumption. / 按用户提供的精度记录日期；不知道时留空，完全没说电影观影日期时才默认今天并说明。
- Record documentary episodes without claiming the whole season was finished. If the volume is unclear, preserve that uncertainty. / 纪录片只记录实际看过的集数，卷数不明就标注待确认。
- The owner can name the active workbook `观影记录-YYMMDD.xlsx`, where the suffix is its last modification date. Writers accept an explicit output path or an authorized `--in-place` update. / 文件名可以显示最后修改日期；更新同一天的文件可以原位写入。

## Install

```text
codex plugin marketplace add wanzhonggu473-kuuu/movie-to-excel --ref main
```

In the Codex plugin directory, select the **Movie to Excel** marketplace and install **movie-to-excel**. The scripts use the spreadsheet runtime bundled with Codex.

## Use

Examples:

```text
Add A City of Sadness to my movie archive. I watched it on 26-09-05.
```

```text
Add this poster as a movie. I only remember watching it in 2025.
```

```text
Add this TV series: season 1, watched in September 2026.
```

```text
Add this Netflix documentary series. I watched episodes 1–3 from 26-09-28 to 26-09-30, but I do not know the volume.
```

For direct script use, pass verified metadata in a JSON file:

```text
node plugins/movie-to-excel/scripts/add-movie.mjs --input my-archive.xlsx --output updated-archive.xlsx --movie-file verified-movie.json
node plugins/movie-to-excel/scripts/add-series.mjs --type documentary --input my-archive.xlsx --output updated-archive.xlsx --record-file verified-documentary.json
```

Use `--type tv` for a TV series. Use `--update-existing` when adding more episodes or seasons to an existing entry. `--in-place` is required to overwrite the input workbook. For a fresh archive, copy the blank template and choose an output filename. The `add-movie.mjs` writer detects a likely duplicate by title and year; `--allow-duplicate` is reserved for intentional repeat viewings.

## Privacy and development

The repository includes an empty workbook template. Personal archives belong in `outputs/` or another local folder; `outputs/` and `personal-archives/` are ignored by Git. There are no accounts, public profiles, followers, feeds, or cloud-hosted viewing histories. See [SECURITY.md](SECURITY.md) for the data boundary.

Version `0.6.0` covers local movie, TV, and documentary recording. Movie metadata and workbook views are synchronized by the scripts. Poster identification still depends on the assistant's verification, and country/region classification may require a judgment for international co-productions. Contributions are welcome; see [CONTRIBUTING.md](CONTRIBUTING.md). The project uses the [MIT License](LICENSE).

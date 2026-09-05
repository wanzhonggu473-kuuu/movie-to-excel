# Movie to Excel

Movie to Excel is an open-source, local-first Codex plugin that turns a movie title or poster into a private Excel viewing archive. The workbook stays on the user's computer; internet access is used only to identify a movie and verify public metadata.

## What it does

Give Codex a movie title, poster, or screenshot. The plugin resolves the exact film, verifies its Chinese and official English titles, director, year, production region, genre, and source, then creates or updates a local `.xlsx` archive. It detects likely duplicates and keeps personal ratings and notes private.

The included workbook offers three interchangeable views: a two-row region overview, a complete horizontal region layout, and a searchable standard list. A row-based database remains the single source of truth, so presentation can change without losing records.

## Install from GitHub

After this repository is published, replace `<owner>` with the GitHub account or organization that hosts it:

```text
codex plugin marketplace add <owner>/movie-to-excel --ref main
```

Then open the Plugins Directory in the ChatGPT desktop app, choose the **Movie to Excel** marketplace, and install **movie-to-excel**. GitHub marketplace sources and the required repository layout follow the official [OpenAI plugin packaging documentation](https://developers.openai.com/plugins/build/plugins).

## Use

Examples of natural requests:

```text
Add A City of Sadness to my movie archive. I watched it today.
```

```text
Identify this poster and add the movie to my existing Excel archive. My rating is 8.5.
```

If a title is ambiguous, the plugin asks for the year or version before writing. If no viewing date is supplied, it uses the user's current local date and states that assumption.

## Workbook structure

- `我的电影`: fixed-height overview, with Mainland China, Hong Kong, and Taiwan on the first row and the United States, Europe/Oceania, and other regions on the second.
- `地区横向`: complete horizontal region layout.
- `标准清单`: searchable and sortable list.
- `电影数据库`: canonical row-based data source.
- `统计`: local totals by region.
- `设置`: presentation, language, date, and lookup preferences.

Titles appear as `中文名 / Official English Title` in grouped views. Directors use established Chinese names for productions from Mainland China, Hong Kong, or Taiwan; all other productions use full English or romanized names.

## Privacy

There are no accounts, public profiles, followers, likes, feeds, analytics, or cloud-hosted viewing histories. Ratings and private notes remain inside the local workbook. See [SECURITY.md](SECURITY.md) for the data boundary.

## Repository layout

```text
.agents/plugins/marketplace.json
plugins/movie-to-excel/
├── .codex-plugin/plugin.json
├── assets/movie-archive-template.xlsx
├── examples/verified-movie.example.json
├── scripts/
└── skills/movie-to-excel/
```

## Development status

Version `0.5.0` is an early working prototype. Title-based identification, region classification, duplicate detection, workbook updates, and synchronized views have been exercised end to end. Poster-identification tests, configurable region rules, larger database capacity, and easier non-Codex installation remain on the roadmap.

## Contributing and license

Contributions are welcome; read [CONTRIBUTING.md](CONTRIBUTING.md) before opening a pull request. Movie to Excel is released under the [MIT License](LICENSE).

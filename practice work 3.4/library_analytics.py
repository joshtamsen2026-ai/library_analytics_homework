"""
Library Analytics
=================

"What":  Runs five analytical SQL queries against a small library database.
"Why":   Turns raw checkout data into insights about members, genres, and books.
"How":   Uses JOIN, GROUP BY, HAVING, COUNT, AVG, and subqueries in SQLite, then:
           1. Prints clearly labeled results to the console.
           2. Exports all results to a professional CSV file.
           3. Builds a DOCX report with tables, charts, and a findings summary.

Usage:
    python library_analytics.py            # "Run, save, and open both reports"
    python library_analytics.py --no-open  # "Run and save without opening files"
    python library_analytics.py --force    # "Rebuild reports even if unchanged"
"""

## ===========================================================================
## STEP 1: IMPORTS
## "Why":  Standard-library modules handle the database, files, hashing, and
##         dates. They load almost instantly.
## "How":  The heavy third-party libraries (matplotlib ~0.8 s, python-docx
##         ~0.1 s) are imported LAZILY inside the chart and DOCX functions.
##         They are only loaded when the reports actually need rebuilding,
##         which is the main reason the program responds so quickly.
## ===========================================================================
import csv
import hashlib
import io
import json
import os
import sqlite3
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

## ===========================================================================
## STEP 2: CONFIGURATION CONSTANTS
## "Why":  Keeping file paths and colors in one place makes them easy to change
##         without searching through the whole file.
## "How":  Output files are saved next to this script, no matter which folder
##         the script is run from.
## ===========================================================================
Row = tuple[Any, ...]
Results = dict[str, list[Row]]

OUTPUT_DIR: Path = Path(__file__).resolve().parent
CSV_PATH: Path = OUTPUT_DIR / "library_analytics_results.csv"
DOCX_PATH: Path = OUTPUT_DIR / "library_analytics_report.docx"

## "Why": Stores a fingerprint of the last build so unchanged reports are reused.
CACHE_PATH: Path = OUTPUT_DIR / ".library_analytics_cache.json"

## "Why": Chart quality settings. 150 dpi is still print quality but renders
##        and compresses much faster than 200 dpi.
CHART_DPI: int = 150
PNG_COMPRESS_LEVEL: int = 1

## "Why": Stored as hex text so python-docx is not needed at import time.
BRAND_COLOR_HEX: str = "1F3A5F"
CHART_COLORS: list[str] = ["#1F3A5F", "#3D7EAA", "#7FB3D5", "#F4A259", "#BC4B51"]
HIGHLIGHT_COLOR: str = "#F4A259"
NEUTRAL_COLOR: str = "#B0BEC5"
REFERENCE_LINE_COLOR: str = "#BC4B51"
TITLE_COLOR: str = "#1F3A5F"

HEAVY_BORROWER_THRESHOLD: int = 2


## ===========================================================================
## STEP 3: DATABASE SETUP
## "Why":  The exercise needs three tables (members, books, checkouts) with
##         fixed sample data so every run gives the same results.
## "How":  An in-memory SQLite database (":memory:") is created fresh on every
##         run, so no database file is left on disk.
##         NOTE: Do not change the schema or sample data (exercise rule).
## ===========================================================================
SCHEMA_SQL: str = """
    CREATE TABLE members (
        id   INTEGER PRIMARY KEY,
        name TEXT    NOT NULL,
        tier TEXT    NOT NULL CHECK (tier IN ('standard', 'premium'))
    );

    CREATE TABLE books (
        id    INTEGER PRIMARY KEY,
        title TEXT    NOT NULL,
        genre TEXT    NOT NULL,
        pages INTEGER NOT NULL
    );

    CREATE TABLE checkouts (
        id            INTEGER PRIMARY KEY,
        member_id     INTEGER NOT NULL REFERENCES members(id),
        book_id       INTEGER NOT NULL REFERENCES books(id),
        checkout_date TEXT    NOT NULL,
        return_date   TEXT              -- "NULL means not yet returned"
    );
"""

## "What": 20 members. The original 5 starter members (ids 1-5), plus
##         15 extra members (ids 6-20) with randomly chosen names and tiers.
## "Why":  The extra members have no checkouts yet, so they show how the
##         LEFT JOIN queries handle inactive members and how they pull down
##         the average used by the subquery in Query 5.
MEMBERS_DATA: list[Row] = [
    (1, "Alice", "premium"),
    (2, "Bob", "standard"),
    (3, "Carmen", "premium"),
    (4, "Dev", "standard"),
    (5, "Ella", "standard"),
    (6, "Farid", "standard"),
    (7, "Grace", "premium"),
    (8, "Hiro", "standard"),
    (9, "Isla", "standard"),
    (10, "Jamal", "premium"),
    (11, "Keiko", "standard"),
    (12, "Liam", "standard"),
    (13, "Maya", "premium"),
    (14, "Nikolai", "standard"),
    (15, "Olivia", "standard"),
    (16, "Priya", "premium"),
    (17, "Quinn", "standard"),
    (18, "Rosa", "standard"),
    (19, "Santiago", "premium"),
    (20, "Tara", "standard"),
]

## "What": 10 books across 3 genres (Fantasy, Sci-Fi, Classic).
BOOKS_DATA: list[Row] = [
    (1, "The Hobbit", "Fantasy", 310),
    (2, "A Wizard of Earthsea", "Fantasy", 183),
    (3, "The Name of the Wind", "Fantasy", 662),
    (4, "Dune", "Sci-Fi", 412),
    (5, "Neuromancer", "Sci-Fi", 271),
    (6, "Foundation", "Sci-Fi", 255),
    (7, "The Left Hand of Darkness", "Sci-Fi", 304),
    (8, "Pride and Prejudice", "Classic", 279),
    (9, "Moby-Dick", "Classic", 635),
    (10, "Jane Eyre", "Classic", 507),
]

## "What": 12 checkouts. "None" becomes SQL NULL (book not yet returned).
CHECKOUTS_DATA: list[Row] = [
    (1, 1, 1, "2024-01-05", "2024-01-19"),
    (2, 1, 4, "2024-01-20", "2024-02-03"),
    (3, 1, 3, "2024-02-10", None),
    (4, 1, 8, "2024-02-15", "2024-03-01"),
    (5, 2, 4, "2024-01-08", "2024-01-22"),
    (6, 2, 5, "2024-02-01", "2024-02-14"),
    (7, 2, 2, "2024-02-20", None),
    (8, 3, 9, "2024-01-12", "2024-02-02"),
    (9, 3, 1, "2024-02-05", "2024-02-18"),
    (10, 4, 6, "2024-01-15", "2024-01-29"),
    (11, 4, 7, "2024-02-12", None),
    (12, 5, 10, "2024-02-08", "2024-02-25"),
]


def create_database() -> sqlite3.Connection:
    """Create the in-memory database, build the schema, and load sample data."""
    ## "Why": executescript() runs several CREATE TABLE statements at once.
    conn = sqlite3.connect(":memory:")
    conn.executescript(SCHEMA_SQL)

    ## "Why": executemany() with "?" placeholders inserts many rows safely
    ##        (parameterized queries prevent SQL injection).
    conn.executemany(
        "INSERT INTO members (id, name, tier) VALUES (?, ?, ?)",
        MEMBERS_DATA,
    )
    conn.executemany(
        "INSERT INTO books (id, title, genre, pages) VALUES (?, ?, ?, ?)",
        BOOKS_DATA,
    )
    conn.executemany(
        "INSERT INTO checkouts "
        "(id, member_id, book_id, checkout_date, return_date) "
        "VALUES (?, ?, ?, ?, ?)",
        CHECKOUTS_DATA,
    )
    conn.commit()
    return conn


## ===========================================================================
## STEP 4: ANALYTICAL QUERIES
## "Why":  Each dictionary holds everything about one query (SQL, title,
##         question, column headers). The console, CSV, and DOCX outputs all
##         loop over this single list, so they always stay in sync.
## "How":  To add a new query, add one dictionary here. Every report picks
##         it up automatically.
## ===========================================================================

## "Why": This subquery calculates each member's checkout count. LEFT JOIN keeps
##        members with zero checkouts, so the average stays accurate.
MEMBER_COUNTS_SUBQUERY: str = """
    SELECT COUNT(c2.id) AS cnt
    FROM members m2
    LEFT JOIN checkouts c2 ON c2.member_id = m2.id
    GROUP BY m2.id
"""

QUERIES: list[dict[str, Any]] = [
    ## -----------------------------------------------------------------------
    ## QUERY 1: "How many checkouts does each member have, most active first?"
    ## "How": LEFT JOIN + GROUP BY + COUNT. LEFT JOIN keeps inactive members.
    ##        "Not Yet Returned" counts rows where return_date IS NULL.
    ##        Ties are sorted by name so the output order is predictable.
    ## -----------------------------------------------------------------------
    {
        "key": "member_checkouts",
        "title": "1. Checkouts per Member (Most Active First)",
        "technique": "JOIN + GROUP BY + COUNT",
        "question": "How many checkouts does each member have, most active first?",
        "sql": """
            SELECT m.name      AS member,
                   m.tier      AS tier,
                   COUNT(c.id) AS checkouts,
                   SUM(CASE WHEN c.id IS NOT NULL AND c.return_date IS NULL
                            THEN 1 ELSE 0 END) AS not_returned
            FROM members m
            LEFT JOIN checkouts c ON c.member_id = m.id
            GROUP BY m.id, m.name, m.tier
            ORDER BY checkouts DESC, m.name
        """,
        "headers": ["Member", "Tier", "Checkouts", "Not Yet Returned"],
    },
    ## -----------------------------------------------------------------------
    ## QUERY 2: "Which genres are checked out most often?"
    ## "How": JOIN checkouts to books to find each loan's genre, then COUNT.
    ##        COUNT(DISTINCT ...) shows how broad the demand is.
    ## -----------------------------------------------------------------------
    {
        "key": "genre_popularity",
        "title": "2. Most Checked-Out Genres",
        "technique": "JOIN + GROUP BY + COUNT",
        "question": "Which genres are checked out most often?",
        "sql": """
            SELECT b.genre                     AS genre,
                   COUNT(c.id)                 AS checkouts,
                   COUNT(DISTINCT c.book_id)   AS distinct_titles,
                   COUNT(DISTINCT c.member_id) AS distinct_borrowers
            FROM checkouts c
            JOIN books b ON b.id = c.book_id
            GROUP BY b.genre
            ORDER BY checkouts DESC, b.genre
        """,
        "headers": [
            "Genre",
            "Checkouts",
            "Distinct Titles Borrowed",
            "Distinct Borrowers",
        ],
    },
    ## -----------------------------------------------------------------------
    ## QUERY 3: "Which members have more than 2 checkouts?"
    ## "Why HAVING (not WHERE)": WHERE filters rows BEFORE grouping, and
    ##        HAVING filters groups AFTER COUNT() is calculated.
    ## -----------------------------------------------------------------------
    {
        "key": "heavy_members",
        "title": "3. Members with More Than 2 Checkouts",
        "technique": "GROUP BY + HAVING",
        "question": "Which members have more than 2 checkouts?",
        "sql": f"""
            SELECT m.name      AS member,
                   m.tier      AS tier,
                   COUNT(c.id) AS checkouts
            FROM members m
            JOIN checkouts c ON c.member_id = m.id
            GROUP BY m.id, m.name, m.tier
            HAVING COUNT(c.id) > {HEAVY_BORROWER_THRESHOLD}
            ORDER BY checkouts DESC, m.name
        """,
        "headers": ["Member", "Tier", "Checkouts"],
    },
    ## -----------------------------------------------------------------------
    ## QUERY 4: "What is the average book length (pages) per genre?"
    ## "How": AVG gives the typical length. MIN and MAX show the range.
    ##        ROUND keeps the output tidy (1 decimal place).
    ## -----------------------------------------------------------------------
    {
        "key": "avg_pages",
        "title": "4. Average Book Length per Genre",
        "technique": "GROUP BY + AVG",
        "question": "What is the average book length (pages) per genre?",
        "sql": """
            SELECT genre                AS genre,
                   COUNT(*)             AS books,
                   ROUND(AVG(pages), 1) AS avg_pages,
                   MIN(pages)           AS shortest,
                   MAX(pages)           AS longest
            FROM books
            GROUP BY genre
            ORDER BY avg_pages DESC
        """,
        "headers": ["Genre", "Books", "Average Pages", "Shortest", "Longest"],
    },
    ## -----------------------------------------------------------------------
    ## QUERY 5: "Which members have an above-average number of checkouts?"
    ## "How": The inner subquery counts checkouts per member. The outer
    ##        subquery averages those counts. HAVING keeps only members whose
    ##        count is greater than that average.
    ## -----------------------------------------------------------------------
    {
        "key": "above_average",
        "title": "5. Members with Above-Average Checkouts",
        "technique": "Subquery (compare each member to the overall average)",
        "question": "Which members have an above-average number of checkouts?",
        "sql": f"""
            SELECT m.name      AS member,
                   m.tier      AS tier,
                   COUNT(c.id) AS checkouts,
                   (SELECT ROUND(AVG(cnt), 2)
                    FROM ({MEMBER_COUNTS_SUBQUERY})) AS member_average
            FROM members m
            LEFT JOIN checkouts c ON c.member_id = m.id
            GROUP BY m.id, m.name, m.tier
            HAVING COUNT(c.id) > (
                SELECT AVG(cnt) FROM ({MEMBER_COUNTS_SUBQUERY})
            )
            ORDER BY checkouts DESC, m.name
        """,
        "headers": ["Member", "Tier", "Checkouts", "Average Across All Members"],
    },
]


def run_queries(conn: sqlite3.Connection) -> Results:
    """Run every query in QUERIES and return results keyed by query "key"."""
    return {query["key"]: conn.execute(query["sql"]).fetchall() for query in QUERIES}


def dataset_overview(conn: sqlite3.Connection) -> dict[str, Any]:
    """Return high-level dataset metrics used in the overview sections."""

    ## "Why": A small helper avoids repeating ".execute(...).fetchone()[0]".
    def scalar(sql: str) -> Any:
        return conn.execute(sql).fetchone()[0]

    first_date = scalar("SELECT MIN(checkout_date) FROM checkouts")
    last_date = scalar("SELECT MAX(checkout_date) FROM checkouts")

    return {
        "Members": scalar("SELECT COUNT(*) FROM members"),
        "Premium Members": scalar(
            "SELECT COUNT(*) FROM members WHERE tier = 'premium'"
        ),
        "Standard Members": scalar(
            "SELECT COUNT(*) FROM members WHERE tier = 'standard'"
        ),
        "Books": scalar("SELECT COUNT(*) FROM books"),
        "Genres": scalar("SELECT COUNT(DISTINCT genre) FROM books"),
        "Total Checkouts": scalar("SELECT COUNT(*) FROM checkouts"),
        "Returned": scalar(
            "SELECT COUNT(*) FROM checkouts WHERE return_date IS NOT NULL"
        ),
        "Not Yet Returned": scalar(
            "SELECT COUNT(*) FROM checkouts WHERE return_date IS NULL"
        ),
        "Checkout Period": f"{first_date} to {last_date}",
        "Never-Borrowed Books": scalar(
            "SELECT COUNT(*) FROM books WHERE id NOT IN (SELECT book_id FROM checkouts)"
        ),
    }


## ===========================================================================
## STEP 5: FINDINGS AND RECOMMENDATIONS
## "Why":  The written explanations are built FROM the query results instead
##         of being hard-coded, so the text stays accurate if the data changes.
## "How":  Each finding reads values from the result rows (by column position)
##         and puts them into a readable sentence.
## ===========================================================================
def build_findings(results: Results, overview: dict[str, Any]) -> dict[str, str]:
    """Create plain-English findings for each analysis category."""
    members = results["member_checkouts"]
    genres = results["genre_popularity"]
    heavy = results["heavy_members"]
    pages = results["avg_pages"]
    above = results["above_average"]
    total_checkouts = overview["Total Checkouts"]

    ## "Why": Results are already sorted, so first/last rows are the extremes.
    top_member, least_member = members[0], members[-1]
    top_genre, low_genre = genres[0], genres[-1]
    longest_genre, shortest_genre = pages[0], pages[-1]

    ## "Why": If nobody is above average, calculate the average directly.
    average = above[0][3] if above else round(total_checkouts / overview["Members"], 2)
    top_genre_share = top_genre[1] / total_checkouts * 100
    above_share = sum(row[2] for row in above) / total_checkouts * 100
    return_rate = overview["Returned"] / total_checkouts * 100

    ## "How": Group checkout counts by tier, then average them per tier.
    tier_counts: dict[str, list[int]] = {}
    ## "Note": Each row is (member, tier, checkouts, not_returned).
    for row in members:
        tier_counts.setdefault(row[1], []).append(row[2])
    tier_averages = {tier: sum(c) / len(c) for tier, c in tier_counts.items()}

    ## "Why": Several members can tie for the fewest checkouts (e.g. 0), so
    ##        report how many are tied instead of naming just one of them.
    fewest_count = least_member[2]
    fewest_tied = sum(1 for row in members if row[2] == fewest_count)
    if fewest_tied > 1:
        fewest_note = f"{fewest_tied} members are tied for the fewest ({fewest_count})"
    else:
        fewest_note = f"{least_member[0]} has the fewest ({fewest_count})"

    heavy_names = ", ".join(row[0] for row in heavy) or "none"
    above_names = ", ".join(row[0] for row in above) or "none"

    ## "Why": Only claim a length-vs-demand link when the data supports it.
    if longest_genre[0] == low_genre[0]:
        length_note = (
            "The genre with the longest books is also the least borrowed, "
            "suggesting book length may modestly reduce demand."
        )
    else:
        length_note = "Book length shows no clear link to demand in this dataset."

    return {
        "Member Activity": (
            f"{top_member[0]} ({top_member[1]}) is the most active member with "
            f"{top_member[2]} checkouts, while {fewest_note}. Premium members average "
            f"{tier_averages.get('premium', 0):.1f} checkouts versus "
            f"{tier_averages.get('standard', 0):.1f} for standard members."
        ),
        "Genre Demand": (
            f"{top_genre[0]} leads demand with {top_genre[1]} checkouts "
            f"({top_genre_share:.0f}% of all activity), and {low_genre[0]} "
            f"trails with {low_genre[1]}. Demand is relatively balanced across "
            f"the {overview['Genres']} genres, so no genre is being neglected."
        ),
        "High-Volume Borrowers (HAVING)": (
            f"{len(heavy)} member(s) exceed {HEAVY_BORROWER_THRESHOLD} "
            f"checkouts: {heavy_names}. HAVING filters groups after "
            f"aggregation, which WHERE cannot do."
        ),
        "Book Length by Genre (AVG)": (
            f"{longest_genre[0]} titles are the longest on average "
            f"({longest_genre[2]} pages), while {shortest_genre[0]} titles are "
            f"the shortest ({shortest_genre[2]} pages). {length_note}"
        ),
        "Above-Average Members (Subquery)": (
            f"The average member has {average} checkouts. {len(above)} member(s) "
            f"are above average ({above_names}) and account for "
            f"{above_share:.0f}% of all checkouts, showing activity is "
            f"concentrated among a few readers."
        ),
        "Circulation Health": (
            f"{overview['Returned']} of {total_checkouts} checkouts have been "
            f"returned ({return_rate:.0f}% return rate); "
            f"{overview['Not Yet Returned']} items are still out. "
            f"{overview['Never-Borrowed Books']} book(s) have never been borrowed."
        ),
    }


RECOMMENDATIONS: list[str] = [
    "Reward highly active members (e.g., loyalty perks) to sustain engagement.",
    (
        "Run outreach or reading challenges for low-activity members to raise "
        "the average."
    ),
    "Use genre demand data to prioritize new acquisitions and extra copies.",
    (
        "Promote rarely borrowed titles through displays or curated "
        "recommendation lists."
    ),
    ("Follow up on unreturned items to keep inventory available for other members."),
]

## "Why": One short explanation per query, shown under each DOCX section.
SECTION_INSIGHTS: dict[str, str] = {
    "member_checkouts": (
        "Members are ranked by total checkouts using a LEFT JOIN so members with "
        "zero activity would still appear. The 'Not Yet Returned' column "
        "highlights items still on loan."
    ),
    "genre_popularity": (
        "Checkouts are joined to books to attribute each loan to a genre. "
        "Distinct titles and borrowers show whether demand is broad or driven "
        "by a few books or readers."
    ),
    "heavy_members": (
        "HAVING filters aggregated groups. Only members whose checkout count is "
        f"strictly greater than {HEAVY_BORROWER_THRESHOLD} pass the filter "
        "(highlighted in the chart)."
    ),
    "avg_pages": (
        "AVG summarizes the typical book length per genre; MIN and MAX show the "
        "spread, which the chart's range bars visualize."
    ),
    "above_average": (
        "A nested subquery first counts checkouts per member, then averages "
        "those counts. Each member's count is compared against that average in "
        "the HAVING clause."
    ),
}

CONCLUSION_TEXT: str = (
    "Library activity is concentrated among a small group of engaged members, "
    "while genre demand is fairly balanced. Book length may play a modest role "
    "in genre popularity. Focusing on engaging lower-activity members and "
    "promoting less-borrowed titles offers the clearest opportunity to grow "
    "overall circulation."
)


## ===========================================================================
## STEP 6: CONSOLE OUTPUT
## "Why":  The exercise requires each query to print clearly labeled results.
## "How":  Column widths are calculated from the longest value in each column
##         so the table lines up neatly.
## ===========================================================================
def print_table(headers: list[str], rows: list[Row]) -> None:
    """Print rows as an aligned, text-based table."""
    widths = [len(header) for header in headers]
    for row in rows:
        widths = [
            max(width, len(str(value)))
            for width, value in zip(widths, row, strict=True)
        ]

    header_line = "  ".join(
        header.ljust(width) for header, width in zip(headers, widths, strict=True)
    )
    print(f"  {header_line}")
    print(f"  {'-' * len(header_line)}")

    if not rows:
        print("  (no results)")
    for row in rows:
        line = "  ".join(
            str(value).ljust(width) for value, width in zip(row, widths, strict=True)
        )
        print(f"  {line}")


def print_results(results: Results) -> None:
    """Print every query's title, technique, and result table."""
    divider = "=" * 70
    for query in QUERIES:
        print(divider)
        print(query["title"])
        print(f"Technique: {query['technique']}")
        print(divider)
        print_table(query["headers"], results[query["key"]])
        print()


## ===========================================================================
## STEP 7: CSV EXPORT
## "Why":  A CSV opens in Excel or Google Sheets for further analysis.
## "How":  The file is split into labeled sections (overview, each query,
##         findings, recommendations) with blank rows between them.
##         "utf-8-sig" adds a BOM so Excel reads special characters correctly.
##         newline="" stops Windows from inserting extra blank lines.
## ===========================================================================
def write_csv(
    results: Results,
    overview: dict[str, Any],
    findings: dict[str, str],
    generated_at: str,
) -> None:
    """Write all results to a sectioned, Excel-friendly CSV file."""
    with open(CSV_PATH, "w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)

        ## "What": Report header.
        writer.writerow(["LIBRARY ANALYTICS REPORT"])
        writer.writerow(["Generated", generated_at])
        writer.writerow(
            ["Source", "In-memory SQLite database (members, books, checkouts)"]
        )
        writer.writerow([])

        ## "What": Dataset overview metrics.
        writer.writerow(["SECTION", "DATASET OVERVIEW"])
        writer.writerow(["Metric", "Value"])
        writer.writerows(overview.items())
        writer.writerow([])

        ## "What": One section per query.
        for query in QUERIES:
            rows = results[query["key"]]
            writer.writerow(["SECTION", query["title"]])
            writer.writerow(["Question", query["question"]])
            writer.writerow(["SQL Technique", query["technique"]])
            writer.writerow(["Rows Returned", len(rows)])
            writer.writerow(query["headers"])
            if rows:
                writer.writerows(rows)
            else:
                writer.writerow(["(no results)"])
            writer.writerow([])

        ## "What": Key findings.
        writer.writerow(["SECTION", "KEY FINDINGS"])
        writer.writerow(["Category", "Finding"])
        writer.writerows(findings.items())
        writer.writerow([])

        ## "What": Numbered recommendations.
        writer.writerow(["SECTION", "RECOMMENDATIONS"])
        for number, recommendation in enumerate(RECOMMENDATIONS, start=1):
            writer.writerow([number, recommendation])


## ===========================================================================
## STEP 8: CHARTS
## "Why":  Charts make patterns easier to see than tables alone.
## "How":  Charts use matplotlib's object-oriented Figure API with the Agg
##         (non-GUI) canvas, so no windows open. Each chart is rendered to an
##         in-memory PNG buffer (no temporary files or disk I/O) and embedded
##         directly in the DOCX.
## ===========================================================================
def _new_chart(figsize: tuple[float, float]) -> tuple[Any, Any]:
    """Create a new figure and axes that render off-screen with Agg."""
    ## "Why": Lazy import. matplotlib is only loaded when charts are needed.
    from matplotlib.backends.backend_agg import FigureCanvasAgg
    from matplotlib.figure import Figure

    fig = Figure(figsize=figsize)
    FigureCanvasAgg(fig)
    return fig, fig.subplots()


def _style_axes(ax: Any, title: str, xlabel: str, ylabel: str) -> None:
    """Apply consistent, clean styling to a chart."""
    ax.set_title(title, fontsize=13, fontweight="bold", color=TITLE_COLOR)
    ax.set_xlabel(xlabel)
    ax.set_ylabel(ylabel)
    ax.spines[["top", "right"]].set_visible(False)
    ax.grid(axis="y", linestyle="--", alpha=0.4)


def _label_bars(ax: Any, bars: Any) -> None:
    """Write each bar's value just above the bar."""
    for bar in bars:
        ax.annotate(
            f"{bar.get_height():g}",
            (bar.get_x() + bar.get_width() / 2, bar.get_height()),
            ha="center",
            va="bottom",
            fontsize=9,
            xytext=(0, 3),
            textcoords="offset points",
        )


def _render_chart(fig: Any) -> io.BytesIO:
    """Render a figure to an in-memory PNG buffer."""
    ## "Why": Fast PNG compression trades a slightly larger image for much
    ##        quicker encoding. The visual quality is identical.
    buffer = io.BytesIO()
    fig.tight_layout()
    fig.savefig(
        buffer,
        format="png",
        dpi=CHART_DPI,
        pil_kwargs={"compress_level": PNG_COMPRESS_LEVEL},
    )
    buffer.seek(0)
    return buffer


def _chart_member_checkouts(results: Results) -> io.BytesIO:
    """Bar chart: checkouts per member with an average reference line."""
    rows = results["member_checkouts"]
    above_rows = results["above_average"]
    above_names = {row[0] for row in above_rows}

    ## "Why": Above-average members are highlighted so they stand out.
    colors = [
        HIGHLIGHT_COLOR if row[0] in above_names else CHART_COLORS[1] for row in rows
    ]

    fig, ax = _new_chart((7, 4))
    bars = ax.bar([row[0] for row in rows], [row[2] for row in rows], color=colors)
    _label_bars(ax, bars)

    if above_rows:
        average = above_rows[0][3]
        ax.axhline(
            average,
            color=REFERENCE_LINE_COLOR,
            linestyle="--",
            label=f"Average ({average})",
        )
        ax.legend()

    _style_axes(ax, "Checkouts per Member", "Member", "Checkouts")
    ## "Why": With 20 members, rotated names keep the x-axis labels readable.
    ax.tick_params(axis="x", labelrotation=45)
    for label in ax.get_xticklabels():
        label.set_horizontalalignment("right")
    return _render_chart(fig)


def _chart_genre_popularity(results: Results) -> io.BytesIO:
    """Pie chart: each genre's share of total checkouts."""
    rows = results["genre_popularity"]

    fig, ax = _new_chart((6, 4))
    ax.pie(
        [row[1] for row in rows],
        labels=[row[0] for row in rows],
        autopct="%1.0f%%",
        colors=CHART_COLORS[: len(rows)],
        startangle=90,
        wedgeprops={"edgecolor": "white"},
        textprops={"fontsize": 10},
    )
    ax.set_title(
        "Share of Checkouts by Genre",
        fontsize=13,
        fontweight="bold",
        color=TITLE_COLOR,
    )
    return _render_chart(fig)


def _chart_heavy_members(results: Results) -> io.BytesIO:
    """Horizontal bar chart: every member compared to the HAVING threshold."""
    ## "Why": Reverse the order so the most active member appears at the top.
    rows = list(reversed(results["member_checkouts"]))
    colors = [
        HIGHLIGHT_COLOR if row[2] > HEAVY_BORROWER_THRESHOLD else NEUTRAL_COLOR
        for row in rows
    ]

    fig, ax = _new_chart((7, 4))
    ax.barh([row[0] for row in rows], [row[2] for row in rows], color=colors)
    ax.axvline(
        HEAVY_BORROWER_THRESHOLD,
        color=REFERENCE_LINE_COLOR,
        linestyle="--",
        label=f"Threshold (> {HEAVY_BORROWER_THRESHOLD})",
    )
    ax.legend()
    ax.set_title(
        "Members vs. HAVING Threshold",
        fontsize=13,
        fontweight="bold",
        color=TITLE_COLOR,
    )
    ax.set_xlabel("Checkouts")
    ax.spines[["top", "right"]].set_visible(False)
    return _render_chart(fig)


def _chart_avg_pages(results: Results) -> io.BytesIO:
    """Bar chart: average pages per genre with min-to-max range bars."""
    rows = results["avg_pages"]
    genres = [row[0] for row in rows]
    averages = [row[2] for row in rows]

    ## "How": Error bars need distances BELOW and ABOVE each average.
    below = [row[2] - row[3] for row in rows]
    above = [row[4] - row[2] for row in rows]

    fig, ax = _new_chart((7, 4))
    bars = ax.bar(genres, averages, color=CHART_COLORS[: len(rows)])
    _label_bars(ax, bars)
    ax.errorbar(
        genres,
        averages,
        yerr=[below, above],
        fmt="none",
        ecolor="#333333",
        capsize=6,
        label="Shortest to Longest range",
    )
    ax.legend()
    _style_axes(ax, "Average Pages per Genre", "Genre", "Pages")
    return _render_chart(fig)


def build_charts(results: Results) -> dict[str, io.BytesIO]:
    """Build all charts and return PNG buffers keyed by query "key"."""
    ## "Note": Query 5 reuses the member chart (its average line), so it has
    ##         no separate chart.
    return {
        "member_checkouts": _chart_member_checkouts(results),
        "genre_popularity": _chart_genre_popularity(results),
        "heavy_members": _chart_heavy_members(results),
        "avg_pages": _chart_avg_pages(results),
    }


## ===========================================================================
## STEP 9: DOCX REPORT
## "Why":  A Word document is easy to share with people who don't read code.
## "How":  python-docx builds the document in order: title -> summary ->
##         overview -> methodology -> detailed results -> findings ->
##         recommendations -> conclusion.
##         python-docx is imported lazily inside these helpers. After the
##         first import, Python reuses the loaded module, so repeated local
##         imports cost almost nothing.
## ===========================================================================
def _brand_color() -> Any:
    """Return the report's brand color as a python-docx RGBColor."""
    from docx.shared import RGBColor

    return RGBColor.from_string(BRAND_COLOR_HEX)


def _add_heading(doc: Any, text: str, level: int) -> None:
    """Add a heading in the report's brand color."""
    heading = doc.add_heading(text, level=level)
    color = _brand_color()
    for run in heading.runs:
        run.font.color.rgb = color


def _add_table(doc: Any, headers: list[str], rows: list[Row]) -> None:
    """Add a styled table with a bold header row."""
    from docx.enum.table import WD_TABLE_ALIGNMENT

    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Light Grid Accent 1"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    ## "What": Header row.
    for cell, header in zip(table.rows[0].cells, headers, strict=True):
        cell.text = str(header)
        cell.paragraphs[0].runs[0].font.bold = True

    ## "Why": Show a placeholder row instead of an empty table.
    if not rows:
        rows = [("(no results)",) + ("",) * (len(headers) - 1)]

    ## "What": Data rows.
    for row in rows:
        for cell, value in zip(table.add_row().cells, row, strict=True):
            cell.text = str(value)

    doc.add_paragraph()


def _add_label_value(doc: Any, label: str, value: str) -> None:
    """Add a "Label: value" line with the label in bold."""
    from docx.shared import Pt

    paragraph = doc.add_paragraph()
    paragraph.add_run(f"{label}: ").bold = True
    paragraph.add_run(value)
    paragraph.paragraph_format.space_after = Pt(2)


def _add_title_page(doc: Any, generated_at: str) -> None:
    """Add the centered report title, subtitle, and generated date."""
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    title = doc.add_heading("Library Analytics Report", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    color = _brand_color()
    for run in title.runs:
        run.font.color.rgb = color

    subtitle = doc.add_paragraph(
        "Checkout Activity, Genre Demand, and Member Engagement"
    )
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.runs[0].italic = True

    meta = doc.add_paragraph(f"Generated: {generated_at}")
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER


def _add_executive_summary(
    doc: Any, results: Results, overview: dict[str, Any]
) -> None:
    """Add a short summary paragraph for quick readers."""
    top_member = results["member_checkouts"][0]
    top_genre = results["genre_popularity"][0]
    above_count = len(results["above_average"])

    _add_heading(doc, "Executive Summary", 1)
    doc.add_paragraph(
        f"This report analyzes {overview['Total Checkouts']} checkouts made by "
        f"{overview['Members']} members across {overview['Books']} books in "
        f"{overview['Genres']} genres. {top_member[0]} is the most active member "
        f"({top_member[2]} checkouts) and {top_genre[0]} is the most popular "
        f"genre ({top_genre[1]} checkouts). {above_count} member(s) borrow more "
        f"than the average reader."
    )


def _add_methodology(doc: Any) -> None:
    """Explain how the results were produced."""
    _add_heading(doc, "Methodology", 1)
    doc.add_paragraph(
        "All results were produced with SQL queries against a SQLite database "
        "containing three tables: members (id, name, tier), books (id, title, "
        "genre, pages), and checkouts (id, member_id, book_id, checkout_date, "
        "return_date). Techniques used include JOIN, GROUP BY, HAVING, COUNT, "
        "AVG, and subqueries."
    )


def _add_detailed_results(
    doc: Any, results: Results, charts: dict[str, io.BytesIO]
) -> None:
    """Add one section per query: details, table, chart, interpretation."""
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Inches

    _add_heading(doc, "Detailed Results", 1)
    for query in QUERIES:
        key = query["key"]
        _add_heading(doc, query["title"], 2)
        _add_label_value(doc, "Question", query["question"])
        _add_label_value(doc, "SQL Technique", query["technique"])
        _add_label_value(doc, "Rows Returned", str(len(results[key])))
        doc.add_paragraph()

        _add_table(doc, query["headers"], results[key])

        if key in charts:
            doc.add_picture(charts[key], width=Inches(5.8))
            doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

        interpretation = doc.add_paragraph()
        interpretation.add_run("Interpretation: ").bold = True
        interpretation.add_run(SECTION_INSIGHTS[key])


def _add_findings_and_conclusion(doc: Any, findings: dict[str, str]) -> None:
    """Add the findings, recommendations, and conclusion sections."""
    doc.add_page_break()
    _add_heading(doc, "Explanation of Findings", 1)
    doc.add_paragraph(
        "The following explains what each result means and how the results "
        "relate to one another."
    )
    for category, text in findings.items():
        _add_heading(doc, category, 2)
        doc.add_paragraph(text)

    _add_heading(doc, "Recommendations", 1)
    for recommendation in RECOMMENDATIONS:
        doc.add_paragraph(recommendation, style="List Bullet")

    _add_heading(doc, "Conclusion", 1)
    doc.add_paragraph(CONCLUSION_TEXT)


def write_docx(
    results: Results,
    overview: dict[str, Any],
    findings: dict[str, str],
    charts: dict[str, io.BytesIO],
    generated_at: str,
) -> None:
    """Build and save the full DOCX report."""
    from docx import Document
    from docx.shared import Pt

    doc = Document()

    ## "Why": Set one default font so the whole document looks consistent.
    normal_style = doc.styles["Normal"]
    normal_style.font.name = "Calibri"
    normal_style.font.size = Pt(11)

    _add_title_page(doc, generated_at)
    _add_executive_summary(doc, results, overview)
    _add_heading(doc, "Dataset Overview", 1)
    _add_table(doc, ["Metric", "Value"], list(overview.items()))
    _add_methodology(doc)
    _add_detailed_results(doc, results, charts)
    _add_findings_and_conclusion(doc, findings)

    doc.save(str(DOCX_PATH))


## ===========================================================================
## STEP 10: BUILD CACHE (SKIP UNCHANGED REPORTS)
## "Why":  Rebuilding charts and the DOCX is the slowest part of the program.
##         If neither the data nor this script has changed, the existing
##         reports are already correct, so they can be reused instantly.
## "How":  A SHA-256 fingerprint is made from the query results, the dataset
##         overview, and this script's source code. It is saved after each
##         build. On the next run, matching fingerprints plus existing report
##         files mean the rebuild is skipped. Use "--force" to always rebuild.
## ===========================================================================
def compute_fingerprint(results: Results, overview: dict[str, Any]) -> str:
    """Return a SHA-256 fingerprint of the data and this script's code."""
    digest = hashlib.sha256()
    digest.update(
        json.dumps(
            {"results": results, "overview": overview},
            sort_keys=True,
            default=str,
        ).encode("utf-8")
    )
    digest.update(Path(__file__).read_bytes())
    return digest.hexdigest()


def reports_are_current(fingerprint: str) -> bool:
    """Return True if both reports exist and match the saved fingerprint."""
    if not (CSV_PATH.exists() and DOCX_PATH.exists() and CACHE_PATH.exists()):
        return False
    try:
        saved = json.loads(CACHE_PATH.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        ## "Why": A missing or corrupted cache file simply forces a rebuild.
        return False
    return bool(saved.get("fingerprint") == fingerprint)


def save_fingerprint(fingerprint: str, generated_at: str) -> None:
    """Save the fingerprint of the reports that were just built."""
    CACHE_PATH.write_text(
        json.dumps({"fingerprint": fingerprint, "generated_at": generated_at}),
        encoding="utf-8",
    )


## ===========================================================================
## STEP 11: OPEN GENERATED FILES
## "Why":  Opening the reports automatically lets the user see them at once.
## "How":  Each operating system has a different "open" command:
##         Windows -> os.startfile, macOS -> "open", Linux -> "xdg-open".
##         subprocess.run with a list is safer than os.system, which passes
##         the command through a shell.
## ===========================================================================
def open_file(path: Path) -> None:
    """Open a file with the operating system's default application."""
    try:
        if sys.platform == "win32":
            ## "Note": This API is available only on Windows.
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)
        else:
            subprocess.run(["xdg-open", str(path)], check=False)
    except OSError as error:
        print(f"Could not open {path}: {error}")


## ===========================================================================
## STEP 12: MAIN PROGRAM FLOW
## "Why":  Keeping the overall flow in one function makes the program easy to
##         follow from top to bottom.
## "How":  1. Build database   2. Run queries   3. Print results
##         4. Skip to step 7 if the reports are already up to date
##         5. Build findings + write CSV        6. Build charts + DOCX
##         7. Open both files (unless "--no-open" is passed)
## ===========================================================================
def main() -> None:
    """Run the full analytics pipeline."""
    ## "Why": try/finally makes sure the database connection always closes,
    ##        even if a query fails.
    conn = create_database()
    try:
        results = run_queries(conn)
        overview = dataset_overview(conn)
    finally:
        conn.close()

    print_results(results)

    fingerprint = compute_fingerprint(results, overview)
    if "--force" not in sys.argv and reports_are_current(fingerprint):
        print("Reports are up to date (data unchanged). Use --force to rebuild.")
    else:
        generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        findings = build_findings(results, overview)
        write_csv(results, overview, findings, generated_at)
        write_docx(results, overview, findings, build_charts(results), generated_at)
        save_fingerprint(fingerprint, generated_at)

    print(f"CSV report:  {CSV_PATH}")
    print(f"DOCX report: {DOCX_PATH}")

    if "--no-open" not in sys.argv:
        open_file(CSV_PATH)
        open_file(DOCX_PATH)


## "Why": This guard runs main() only when the file is executed directly,
##        not when it is imported by another script (for example, tests).
if __name__ == "__main__":
    main()

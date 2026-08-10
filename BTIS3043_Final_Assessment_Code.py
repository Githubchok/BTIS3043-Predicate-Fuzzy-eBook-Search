"""
BTIS3043 2026B Final Assessment
Predicate + Fuzzy eBook Search System

Requirements implemented:
- Query Dataset A, B and C for both fixed scenarios
- Predicate-only filtering
- Fuzzy suitability scoring
- Ranking of fuzzy-enhanced results
- Comparison of predicate-only vs fuzzy-enhanced results
- Scenario 1 relationship classification
- Scenario 2 Current Subscription output
- Dataset size and result summary
- CSV outputs for reproducibility

The script uses pandas and Python standard library.
"""

from pathlib import Path
import math
import pandas as pd


# ============================================================
# 1. FILE SETTINGS
# ============================================================

BASE_DIR = Path(__file__).resolve().parent


# The code supports both the original filenames and filenames
# containing "(1)".
def find_dataset(possible_names):
    for filename in possible_names:
        path = BASE_DIR / filename
        if path.exists():
            return path

    raise FileNotFoundError(
        "Dataset file not found. Tried:\n"
        + "\n".join(str(BASE_DIR / name) for name in possible_names)
    )


DATASET_FILES = {
    "A": find_dataset([
        "BTIS3043_Dataset_A_Existing_eBook_Collection.xlsx",
        "BTIS3043_Dataset_A_Existing_eBook_Collection(1).xlsx"
    ]),

    "B": find_dataset([
        "BTIS3043_Dataset_B_Academic_eBook_Catalogue.xlsx",
        "BTIS3043_Dataset_B_Academic_eBook_Catalogue(1).xlsx"
    ]),

    "C": find_dataset([
        "BTIS3043_Dataset_C_eBook_Acquisition_Catalogue.xlsx",
        "BTIS3043_Dataset_C_eBook_Acquisition_Catalogue(1).xlsx"
    ])
}


OUTPUT_DIR = BASE_DIR / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)


CURRENT_YEAR = 2026


# ============================================================
# 2. FIXED SCENARIOS
# ============================================================

SCENARIOS = {

    # --------------------------------------------------------
    # Scenario 1
    # --------------------------------------------------------

    "Scenario_1_AI_Programming_Mathematics": {

        "topic_keywords": [
            "artificial intelligence",
            "intelligent systems",
            "machine learning",
            "deep learning",
            "computer vision",
            "robotics",
            "expert systems",
            "knowledge representation",
            "data science",
            "analytics",
            "algorithm",
            "algorithms",
            "data structures",
            "software engineering",
            "python",
            "java",
            "c++",
            "programming",
            "statistics",
            "probability",
            "linear algebra",
            "discrete mathematics",
            "calculus",
            "optimization",
            "decision analysis",
            "mathematics"
        ],

        "direct_ai_keywords": [
            "artificial intelligence",
            "intelligent systems",
            "machine learning",
            "deep learning",
            "computer vision",
            "robotics",
            "expert systems",
            "knowledge representation"
        ],

        "programming_keywords": [
            "programming",
            "python",
            "java",
            "c++",
            "algorithm",
            "algorithms",
            "data structures",
            "software engineering",
            "data science"
        ],

        "math_keywords": [
            "statistics",
            "probability",
            "linear algebra",
            "discrete mathematics",
            "calculus",
            "optimization",
            "decision analysis",
            "mathematics"
        ],

        "preferred_formats": [
            "epub",
            "pdf",
            "adobe reader"
        ],

        "top_n": 5
    },


    # --------------------------------------------------------
    # Scenario 2
    # --------------------------------------------------------

    "Scenario_2_Cybersecurity_Secure_Computing": {

        "topic_keywords": [
            "cybersecurity",
            "cyber security",
            "computer security",
            "network security",
            "information security",
            "security",
            "cryptography",
            "privacy",
            "digital forensics",
            "forensics",
            "information assurance",
            "secure systems",
            "secure computing",
            "security in computing"
        ],

        "direct_security_keywords": [
            "cybersecurity",
            "cyber security",
            "computer security",
            "network security",
            "information security",
            "cryptography",
            "digital forensics",
            "information assurance",
            "secure systems",
            "security in computing"
        ],

        "preferred_formats": [
            "epub",
            "pdf",
            "adobe reader"
        ],

        "top_n": 10
    }
}


# ============================================================
# 3. DATA LOADING
# ============================================================

def load_datasets():

    """Load the three teacher-provided Excel datasets."""

    datasets = {}

    for name, path in DATASET_FILES.items():

        if not path.exists():

            raise FileNotFoundError(
                f"Dataset {name} not found:\n{path}"
            )

        datasets[name] = pd.read_excel(path)

    return datasets


# ============================================================
# 4. HELPER FUNCTIONS
# ============================================================

def clean_text(value):

    """Convert a value to searchable lowercase text."""

    if pd.isna(value):
        return ""

    return str(value).strip().lower()


def combined_text(row, dataset_name):

    """
    Build searchable text from fields available
    in each dataset.

    The datasets are intentionally NOT merged.
    """

    if dataset_name == "A":

        fields = [
            "Title",
            "Author",
            "Recommended by"
        ]

    elif dataset_name == "B":

        fields = [
            "Discipline (Level 1)",
            "Discipline (Level 2)",
            "Discipline (Level 3)",
            "Discipline (Level 4)",
            "Title",
            "Author",
            "eBook Format",
            "Origin "
        ]

    else:

        fields = [
            "Category",
            "Discipline",
            "Title",
            "Author",
            "eBook Format"
        ]

    return " ".join(
        clean_text(row.get(field, ""))
        for field in fields
    )


def keyword_hits(text, keywords):

    """Return matched keywords."""

    text = clean_text(text)

    return [
        keyword
        for keyword in keywords
        if keyword.lower() in text
    ]


# ============================================================
# 5. FUZZY MEMBERSHIP FUNCTIONS
# ============================================================

def keyword_relevance(text, keywords):

    """
    Fuzzy membership for general topic relevance.

    0.00 = no relevant keyword
    Higher values = more different relevant matches
    """

    hits = keyword_hits(text, keywords)

    if not hits:
        return 0.0

    return min(1.0, len(hits) / 4.0)


def direct_topic_relevance(text, keywords):

    """
    Stronger fuzzy membership for direct topic relevance.

    A direct AI/security keyword contributes strongly.
    """

    hits = keyword_hits(text, keywords)

    if not hits:
        return 0.0

    return min(1.0, len(hits) / 2.0)


def recency_membership(year):

    """
    Fuzzy membership for recency.

    2026 = 1.00
    2025 = 0.80
    2024 = 0.60
    2023 = 0.40
    2022 = 0.20
    2021 or older = 0.00
    """

    try:
        year = float(year)

    except (TypeError, ValueError):
        return 0.0

    age = CURRENT_YEAR - year

    if age <= 0:
        return 1.0

    if age >= 5:
        return 0.0

    return round(1.0 - (age / 5.0), 3)


def affordability_membership(price):

    """
    Fuzzy affordability membership.

    <= 200  -> 1.00
    >= 800  -> 0.00
    Between -> linear decrease

    Missing price -> 0.50 neutral membership.
    """

    if price is None or pd.isna(price):
        return 0.50

    try:
        price = float(price)

    except (TypeError, ValueError):
        return 0.50

    if price <= 200:
        return 1.0

    if price >= 800:
        return 0.0

    return round((800 - price) / 600, 3)


def format_membership(fmt, preferred_formats):

    """Fuzzy membership for suitable eBook format."""

    fmt = clean_text(fmt)

    if not fmt:
        return 0.50

    if any(
        preferred in fmt
        for preferred in preferred_formats
    ):
        return 1.0

    return 0.50


# ============================================================
# 6. DATASET-SPECIFIC FIELDS
# ============================================================

def get_year(row, dataset_name):

    if dataset_name == "A":

        return row.get(
            "Copyright Year",
            None
        )

    if dataset_name == "B":

        copyright_year = row.get(
            "Copyright",
            None
        )

        if pd.notna(copyright_year):
            return copyright_year

        pub_date = row.get(
            "Pub Date",
            None
        )

        if pd.notna(pub_date):

            try:
                return pd.to_datetime(
                    pub_date
                ).year

            except Exception:
                pass

    return row.get(
        "Copyright Year",
        None
    )


def get_price(row, dataset_name):

    """
    Transparent price selection.

    Dataset A:
        Unit Net Price

    Dataset B:
        No comparable price field

    Dataset C:
        April List Price (USD)
    """

    if dataset_name == "A":

        return row.get(
            "Unit Net Price",
            None
        )

    if dataset_name == "C":

        return row.get(
            "April List Price (USD)",
            None
        )

    return None


def get_title(row):

    return row.get(
        "Title",
        ""
    )


def get_format(row, dataset_name):

    if dataset_name == "A":

        # Dataset A is already an existing eBook collection.
        return "existing ebook"

    return row.get(
        "eBook Format",
        ""
    )


def get_record_type(dataset_name):

    """
    Dataset interpretation for reporting.

    Dataset A:
        Existing/current collection

    Dataset B:
        Academic eBook catalogue

    Dataset C:
        eBook acquisition catalogue
    """

    if dataset_name == "A":
        return "Current / Existing eBook Collection"

    if dataset_name == "B":
        return "Academic eBook Catalogue"

    return "eBook Acquisition Catalogue"


# ============================================================
# 7. SCENARIO 1 RELATIONSHIP CLASSIFICATION
# ============================================================

def classify_scenario1_relationship(
    text,
    dataset_name
):

    """
    Identify the primary relationship required by Scenario 1.

    Categories:
    - Direct AI-related
    - Programming support
    - Mathematical support
    - Other justified relationship
    """

    direct_ai = keyword_hits(
        text,
        SCENARIOS[
            "Scenario_1_AI_Programming_Mathematics"
        ]["direct_ai_keywords"]
    )

    programming = keyword_hits(
        text,
        SCENARIOS[
            "Scenario_1_AI_Programming_Mathematics"
        ]["programming_keywords"]
    )

    mathematics = keyword_hits(
        text,
        SCENARIOS[
            "Scenario_1_AI_Programming_Mathematics"
        ]["math_keywords"]
    )


    supporting = []


    # Direct AI takes priority because it is
    # the most direct relationship.

    if direct_ai:

        primary = "Direct AI-related"

        if programming:
            supporting.append(
                "Programming support"
            )

        if mathematics:
            supporting.append(
                "Mathematical support"
            )

        return (
            primary,
            "; ".join(supporting),
            ", ".join(direct_ai)
        )


    # Programming support

    if programming:

        primary = "Programming support"

        if mathematics:
            supporting.append(
                "Mathematical support"
            )

        return (
            primary,
            "; ".join(supporting),
            ", ".join(programming)
        )


    # Mathematical support

    if mathematics:

        return (
            "Mathematical support",
            "",
            ", ".join(mathematics)
        )


    return (
        "Other justified relationship",
        "",
        ""
    )


# ============================================================
# 8. PREDICATE QUERY
# ============================================================

def predicate_match(
    row,
    dataset_name,
    scenario
):

    """
    Predicate reasoning.

    Returns True / False only.

    Scenario 1:
        At least one relevant AI/programming/
        mathematical topic condition.

    Scenario 2:
        At least one relevant security condition.
    """

    text = combined_text(
        row,
        dataset_name
    )


    # --------------------------------------------------------
    # Scenario 1
    # --------------------------------------------------------

    if scenario == "Scenario_1_AI_Programming_Mathematics":

        hits = keyword_hits(
            text,
            SCENARIOS[scenario]["topic_keywords"]
        )

        return len(hits) > 0


    # --------------------------------------------------------
    # Scenario 2
    # --------------------------------------------------------

    if scenario == "Scenario_2_Cybersecurity_Secure_Computing":

        security_hits = keyword_hits(
            text,
            SCENARIOS[
                scenario
            ]["direct_security_keywords"]
        )


        # Dataset A has Recommended by information.
        # CS/IT plus security evidence is accepted.

        if dataset_name == "A":

            recommender = clean_text(
                row.get(
                    "Recommended by",
                    ""
                )
            )

            cs_support = (
                "cs/it" in recommender
            )

            return (
                len(security_hits) > 0
                or (
                    cs_support
                    and "security" in text
                )
            )


        return len(security_hits) > 0


    return False


# ============================================================
# 9. FUZZY EVALUATION
# ============================================================

def fuzzy_score(
    row,
    dataset_name,
    scenario
):

    """
    Calculate fuzzy suitability.

    Scenario 1:
        Topic relevance       50%
        Recency               20%
        Format suitability    15%
        Affordability         15%

    Scenario 2:
        Security relevance    50%
        Recency               20%
        Format suitability    15%
        Affordability         15%
    """

    cfg = SCENARIOS[scenario]

    text = combined_text(
        row,
        dataset_name
    )


    # --------------------------------------------------------
    # Scenario 1 relevance
    # --------------------------------------------------------

    if scenario == "Scenario_1_AI_Programming_Mathematics":

        topic = keyword_relevance(
            text,
            cfg["topic_keywords"]
        )

        direct_ai = direct_topic_relevance(
            text,
            cfg["direct_ai_keywords"]
        )

        # Direct AI receives stronger importance.

        relevance = min(
            1.0,
            0.65 * topic
            + 0.35 * direct_ai
        )


        # Dataset A has Recommended by instead of
        # detailed discipline fields.

        if dataset_name == "A":

            recommender = clean_text(
                row.get(
                    "Recommended by",
                    ""
                )
            )

            programming_terms = [
                "programming",
                "software",
                "algorithm",
                "data structure",
                "python",
                "java",
                "c++"
            ]

            if (
                "cs/it" in recommender
                and any(
                    term in text
                    for term in programming_terms
                )
            ):

                relevance = max(
                    relevance,
                    0.50
                )


    # --------------------------------------------------------
    # Scenario 2 relevance
    # --------------------------------------------------------

    else:

        relevance = direct_topic_relevance(
            text,
            cfg["direct_security_keywords"]
        )


    # --------------------------------------------------------
    # Recency
    # --------------------------------------------------------

    year = get_year(
        row,
        dataset_name
    )

    recency = recency_membership(
        year
    )


    # --------------------------------------------------------
    # Affordability
    # --------------------------------------------------------

    price = get_price(
        row,
        dataset_name
    )

    affordability = affordability_membership(
        price
    )


    # --------------------------------------------------------
    # Format
    # --------------------------------------------------------

    fmt = get_format(
        row,
        dataset_name
    )

    format_score = format_membership(
        fmt,
        cfg["preferred_formats"]
    )


    # --------------------------------------------------------
    # Weighted fuzzy aggregation
    # --------------------------------------------------------

    final_score = (
        0.50 * relevance
        + 0.20 * recency
        + 0.15 * format_score
        + 0.15 * affordability
    )


    return {

        "Relevance":
            round(relevance, 3),

        "Recency":
            round(recency, 3),

        "Format_Suitability":
            round(format_score, 3),

        "Affordability":
            round(affordability, 3),

        "Fuzzy_Score":
            round(final_score, 3),

        "Matched_Keywords":
            ", ".join(
                keyword_hits(
                    text,
                    cfg["topic_keywords"]
                )
            )
    }


# ============================================================
# 10. ADD RESULT FIELDS
# ============================================================

def add_common_result_fields(
    df,
    dataset_name,
    scenario
):

    records = []

    for _, row in df.iterrows():

        result = row.to_dict()


        # Basic identification

        result["Dataset"] = dataset_name

        result["Scenario"] = scenario

        result["Record_Type"] = get_record_type(
            dataset_name
        )


        # Predicate

        result["Predicate_Pass"] = predicate_match(
            row,
            dataset_name,
            scenario
        )


        # Fuzzy evaluation

        fuzzy = fuzzy_score(
            row,
            dataset_name,
            scenario
        )

        result.update(fuzzy)


        # Dataset-specific values

        result["Year_Used"] = get_year(
            row,
            dataset_name
        )

        result["Price_Used"] = get_price(
            row,
            dataset_name
        )


        # ----------------------------------------------------
        # Scenario 1 relationship
        # ----------------------------------------------------

        if scenario == "Scenario_1_AI_Programming_Mathematics":

            text = combined_text(
                row,
                dataset_name
            )

            (
                primary_relationship,
                supporting_relationships,
                relationship_evidence
            ) = classify_scenario1_relationship(
                text,
                dataset_name
            )

            result[
                "Primary_Relationship"
            ] = primary_relationship

            result[
                "Supporting_Relationships"
            ] = supporting_relationships

            result[
                "Relationship_Evidence"
            ] = relationship_evidence


        # ----------------------------------------------------
        # Scenario 2 Current Subscription field
        # ----------------------------------------------------

        if scenario == "Scenario_2_Cybersecurity_Secure_Computing":

            if dataset_name == "A":

                result[
                    "Current_Subscription"
                ] = "Yes - Existing Collection"

            else:

                result[
                    "Current_Subscription"
                ] = "No - Catalogue / Acquisition"


        records.append(result)


    return pd.DataFrame(records)


# ============================================================
# 11. RUN ONE DATASET
# ============================================================

def run_one_dataset(
    df,
    dataset_name,
    scenario
):

    """
    Run predicate + fuzzy reasoning
    for one dataset and scenario.
    """

    all_results = add_common_result_fields(
        df,
        dataset_name,
        scenario
    )


    # --------------------------------------------------------
    # Predicate-only results
    # --------------------------------------------------------

    predicate_results = all_results[
        all_results["Predicate_Pass"] == True
    ].copy()


    # Predicate-only order remains original
    # dataset order.
    predicate_results.insert(
        0,
        "Predicate_Position",
        range(
            1,
            len(predicate_results) + 1
        )
    )


    # --------------------------------------------------------
    # Fuzzy-enhanced results
    # --------------------------------------------------------

    top_n = SCENARIOS[
        scenario
    ]["top_n"]


    fuzzy_results = predicate_results.sort_values(
        by=[
            "Fuzzy_Score",
            "Relevance",
            "Recency"
        ],
        ascending=False
    ).head(top_n).copy()


    fuzzy_results.insert(
        0,
        "Fuzzy_Rank",
        range(
            1,
            len(fuzzy_results) + 1
        )
    )


    return (
        all_results,
        predicate_results,
        fuzzy_results
    )


# ============================================================
# 12. COMPARISON
# ============================================================

def make_comparison(
    predicate_results,
    fuzzy_results
):

    """
    Compare predicate-only position with
    fuzzy-enhanced ranking.
    """

    if predicate_results.empty:

        return pd.DataFrame(
            columns=[
                "Title",
                "Predicate_Position",
                "Fuzzy_Rank",
                "Rank_Change",
                "Fuzzy_Score",
                "Main_Reason"
            ]
        )


    pred = (
        predicate_results
        .copy()
        .reset_index(drop=True)
    )

    fuzzy = fuzzy_results.copy()


    pred_titles = (
        pred["Title"]
        .astype(str)
    )

    rows = []


    for _, frow in fuzzy.iterrows():

        title = str(
            frow["Title"]
        )


        matches = pred[
            pred_titles == title
        ]


        if matches.empty:

            old_position = None
            rank_change = None

        else:

            old_position = int(
                matches.iloc[0][
                    "Predicate_Position"
                ]
            )

            rank_change = (
                old_position
                - int(
                    frow["Fuzzy_Rank"]
                )
            )


        reasons = []


        # Relevance

        if frow["Relevance"] >= 0.75:

            reasons.append(
                "high topic relevance"
            )

        elif frow["Relevance"] >= 0.40:

            reasons.append(
                "moderate topic relevance"
            )


        # Recency

        if frow["Recency"] >= 0.80:

            reasons.append(
                "recent publication"
            )

        elif frow["Recency"] <= 0.20:

            reasons.append(
                "older publication"
            )


        # Affordability

        if frow["Affordability"] >= 0.75:

            reasons.append(
                "relatively affordable"
            )

        elif frow["Affordability"] <= 0.25:

            reasons.append(
                "relatively expensive"
            )


        # Format

        if frow["Format_Suitability"] >= 0.90:

            reasons.append(
                "suitable eBook format"
            )


        rows.append({

            "Title":
                title,

            "Predicate_Position":
                old_position,

            "Fuzzy_Rank":
                int(
                    frow["Fuzzy_Rank"]
                ),

            "Rank_Change":
                rank_change,

            "Fuzzy_Score":
                frow["Fuzzy_Score"],

            "Main_Reason":
                "; ".join(reasons)

        })


    return pd.DataFrame(rows)


# ============================================================
# 13. CURRENT SUBSCRIPTION OUTPUT
# ============================================================

def make_current_subscription_output(
    predicate_results,
    dataset_name,
    scenario
):

    """
    Scenario 2 requirement:

    Display all relevant Current Subscription
    records.

    Dataset A is the Existing eBook Collection,
    so it is treated as the current/existing
    collection for this output.

    Other datasets are not current subscriptions.
    """

    if scenario != "Scenario_2_Cybersecurity_Secure_Computing":

        return pd.DataFrame()


    if dataset_name != "A":

        return pd.DataFrame()


    if predicate_results.empty:

        return pd.DataFrame(
            columns=[
                "Title",
                "Current_Subscription",
                "Predicate_Pass",
                "Fuzzy_Score",
                "Relevance",
                "Recency",
                "Affordability"
            ]
        )


    output = predicate_results.copy()


    columns = [
        "Title",
        "Author",
        "Edition",
        "Copyright Year",
        "Quantity",
        "Recommended by",
        "Unit Net Price",
        "Current_Subscription",
        "Predicate_Pass",
        "Relevance",
        "Recency",
        "Affordability",
        "Fuzzy_Score",
        "Matched_Keywords"
    ]


    available_columns = [
        col
        for col in columns
        if col in output.columns
    ]


    return output[
        available_columns
    ]


# ============================================================
# 14. MAIN PROGRAM
# ============================================================

def main():

    print("=" * 70)

    print(
        "BTIS3043 2026B - "
        "Predicate + Fuzzy eBook Search System"
    )

    print("=" * 70)


    # --------------------------------------------------------
    # Load datasets
    # --------------------------------------------------------

    datasets = load_datasets()


    print("\nDataset sizes:")

    for name, df in datasets.items():

        print(
            f"  Dataset {name}: "
            f"{len(df)} records"
        )


    summary_rows = []


    # --------------------------------------------------------
    # Run both scenarios
    # --------------------------------------------------------

    for scenario in SCENARIOS:

        print("\n" + "=" * 70)

        print(scenario)

        print("=" * 70)


        for dataset_name, df in datasets.items():


            # ------------------------------------------------
            # Run predicate + fuzzy
            # ------------------------------------------------

            (
                all_results,
                predicate_results,
                fuzzy_results
            ) = run_one_dataset(
                df,
                dataset_name,
                scenario
            )


            # ------------------------------------------------
            # Save complete processing output
            # ------------------------------------------------

            all_file = (
                OUTPUT_DIR
                / f"{scenario}_{dataset_name}_all_results.csv"
            )

            all_results.to_csv(
                all_file,
                index=False
            )


            # ------------------------------------------------
            # Save predicate-only output
            # ------------------------------------------------

            pred_file = (
                OUTPUT_DIR
                / f"{scenario}_{dataset_name}_predicate_only.csv"
            )

            predicate_results.to_csv(
                pred_file,
                index=False
            )


            # ------------------------------------------------
            # Save fuzzy-enhanced output
            # ------------------------------------------------

            fuzzy_file = (
                OUTPUT_DIR
                / f"{scenario}_{dataset_name}_fuzzy_enhanced.csv"
            )

            fuzzy_results.to_csv(
                fuzzy_file,
                index=False
            )


            # ------------------------------------------------
            # Save comparison
            # ------------------------------------------------

            comparison = make_comparison(
                predicate_results,
                fuzzy_results
            )


            comparison_file = (
                OUTPUT_DIR
                / f"{scenario}_{dataset_name}_comparison.csv"
            )

            comparison.to_csv(
                comparison_file,
                index=False
            )


            # ------------------------------------------------
            # Scenario 2:
            # Current Subscription output
            # ------------------------------------------------

            current_subscription = (
                make_current_subscription_output(
                    predicate_results,
                    dataset_name,
                    scenario
                )
            )


            if not current_subscription.empty:

                subscription_file = (
                    OUTPUT_DIR
                    / f"{scenario}_{dataset_name}_current_subscription.csv"
                )

                current_subscription.to_csv(
                    subscription_file,
                    index=False
                )


            # ------------------------------------------------
            # Summary
            # ------------------------------------------------

            summary_rows.append({

                "Scenario":
                    scenario,

                "Dataset":
                    dataset_name,

                "Dataset_Size":
                    len(df),

                "Predicate_Matches":
                    len(predicate_results),

                "Fuzzy_Results_Shown":
                    len(fuzzy_results),

                "Top_Fuzzy_Score":
                    (
                        fuzzy_results[
                            "Fuzzy_Score"
                        ].max()
                        if not fuzzy_results.empty
                        else 0
                    )

            })


            # ------------------------------------------------
            # Console output
            # ------------------------------------------------

            print(
                f"\n--- Dataset {dataset_name} ---"
            )

            print(
                f"Records: {len(df)}"
            )

            print(
                f"Predicate matches: "
                f"{len(predicate_results)}"
            )


            # ------------------------------------------------
            # Scenario 2 Current Subscription display
            # ------------------------------------------------

            if (
                scenario
                == "Scenario_2_Cybersecurity_Secure_Computing"
                and dataset_name == "A"
            ):

                print(
                    "\nCurrent Subscription "
                    "/ Existing Collection:"
                )

                if current_subscription.empty:

                    print(
                        "No relevant Current Subscription records."
                    )

                else:

                    display_columns = [
                        col
                        for col in [
                            "Title",
                            "Current_Subscription",
                            "Fuzzy_Score",
                            "Matched_Keywords"
                        ]
                        if col in current_subscription.columns
                    ]

                    print(
                        current_subscription[
                            display_columns
                        ].to_string(
                            index=False
                        )
                    )


            # ------------------------------------------------
            # Fuzzy results
            # ------------------------------------------------

            if fuzzy_results.empty:

                print(
                    "No fuzzy-enhanced results."
                )

            else:

                display_cols = [
                    "Fuzzy_Rank",
                    "Title",
                    "Relevance",
                    "Recency",
                    "Format_Suitability",
                    "Affordability",
                    "Fuzzy_Score",
                    "Matched_Keywords"
                ]


                # Scenario 1 additionally displays
                # relationship classification.

                if (
                    scenario
                    == "Scenario_1_AI_Programming_Mathematics"
                ):

                    display_cols = [
                        "Fuzzy_Rank",
                        "Title",
                        "Primary_Relationship",
                        "Supporting_Relationships",
                        "Relevance",
                        "Recency",
                        "Format_Suitability",
                        "Affordability",
                        "Fuzzy_Score",
                        "Matched_Keywords"
                    ]


                available_display_cols = [
                    col
                    for col in display_cols
                    if col in fuzzy_results.columns
                ]


                print(
                    fuzzy_results[
                        available_display_cols
                    ].to_string(
                        index=False
                    )
                )


    # ========================================================
    # OVERALL SUMMARY
    # ========================================================

    summary = pd.DataFrame(
        summary_rows
    )


    summary.to_csv(
        OUTPUT_DIR / "overall_summary.csv",
        index=False
    )


    # ========================================================
    # DATASET CHARACTERISTICS
    # ========================================================

    characteristics = []


    for dataset_name, df in datasets.items():

        characteristics.append({

            "Dataset":
                dataset_name,

            "Records":
                len(df),

            "Columns":
                len(df.columns),

            "Record_Type":
                get_record_type(
                    dataset_name
                ),

            "Available_Fields":
                ", ".join(
                    str(col)
                    for col in df.columns
                )

        })


    characteristics_df = pd.DataFrame(
        characteristics
    )


    characteristics_df.to_csv(
        OUTPUT_DIR
        / "dataset_characteristics.csv",
        index=False
    )


    # ========================================================
    # FINAL OUTPUT
    # ========================================================

    print("\n" + "=" * 70)

    print("SUMMARY")

    print("=" * 70)

    print(
        summary.to_string(
            index=False
        )
    )


    print("\n" + "=" * 70)

    print("DONE")

    print(
        f"All CSV outputs are saved in:\n"
        f"{OUTPUT_DIR}"
    )

    print("=" * 70)


# ============================================================
# 15. PROGRAM ENTRY POINT
# ============================================================

if __name__ == "__main__":
    main()
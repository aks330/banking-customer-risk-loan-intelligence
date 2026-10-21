#  LIBRARIES

from pathlib import Path

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import matplotlib.dates as mdates


# ================================================================
# CONFIGURATION
# ================================================================

FILE_PATH = r"C:\Users\USER\Downloads\Banking_Loan_Policy_PowerBI_Ready.xlsx"
SHEET_NAME = "Loan_Applications"

OUTPUT_DIR = Path("python_analysis_output")
OUTPUT_DIR.mkdir(exist_ok=True)

# Consistent, professional chart formatting
plt.rcParams["figure.figsize"] = (11, 6)
plt.rcParams["font.size"] = 10
plt.rcParams["axes.titlesize"] = 13
plt.rcParams["axes.labelsize"] = 10
plt.rcParams["figure.dpi"] = 120


# ================================================================
#  DATA LOADING AND PREPROCESSING
# ================================================================

print("\n" + "=" * 65)
print("NOVATRUST BANK — PYTHON DATA ANALYSIS")
print("=" * 65)

# Load Excel file
file_path = Path(FILE_PATH)

if not file_path.exists():
    raise FileNotFoundError(
        f"Dataset not found: {FILE_PATH}\n"
        "Place the Excel file in the same folder as this Python script."
    )

df = pd.read_excel(
    file_path,
    sheet_name=SHEET_NAME
)

# Standardize column names
df.columns = df.columns.astype(str).str.strip()

print("\nOriginal dataset loaded successfully.")
print("Original number of records:", f"{len(df):,}")
print("\nAvailable columns:")
print(df.columns.tolist())

# Required fields for the analysis
required_columns = [
    "Loan_Type",
    "LoanAmount",
    "Policy_Max_Loan",
    "Application_Date",
    "Eligibility_Status"
]

missing_columns = [
    column for column in required_columns
    if column not in df.columns
]

if missing_columns:
    raise ValueError(
        "The dataset is missing required columns: "
        + ", ".join(missing_columns)
    )

# Convert numerical fields
for column in ["LoanAmount", "Policy_Max_Loan"]:
    df[column] = pd.to_numeric(
        df[column],
        errors="coerce"
    )

# Convert application date
df["Application_Date"] = pd.to_datetime(
    df["Application_Date"],
    errors="coerce"
)

# Clean category labels
for column in ["Loan_Type", "Eligibility_Status"]:
    df[column] = (
        df[column]
        .astype("string")
        .str.strip()
    )

# Exclude records missing required fields
before_cleaning = len(df)

analysis_df = df.dropna(
    subset=required_columns
).copy()

# Exclude empty category labels
analysis_df = analysis_df[
    analysis_df["Loan_Type"].ne("")
    & analysis_df["Eligibility_Status"].ne("")
].copy()

# Exclude negative requested loan amounts from monetary analysis.
# Zero amounts are retained, as they may be valid records.
analysis_df = analysis_df[
    analysis_df["LoanAmount"] >= 0
].copy()

after_cleaning = len(analysis_df)
excluded_records = before_cleaning - after_cleaning

# Chronological ordering for time-series analysis
analysis_df = analysis_df.sort_values(
    "Application_Date"
).reset_index(drop=True)

print("\n" + "-" * 65)
print("DATA PREPROCESSING SUMMARY")
print("-" * 65)

print(f"Records before preprocessing : {before_cleaning:,}")
print(f"Records retained             : {after_cleaning:,}")
print(f"Records excluded             : {excluded_records:,}")
print(f"Columns available            : {len(analysis_df.columns)}")

print("\nData types:")
print(
    analysis_df[
        [
            "Loan_Type",
            "LoanAmount",
            "Policy_Max_Loan",
            "Application_Date",
            "Eligibility_Status"
        ]
    ].dtypes
)

if analysis_df.empty:
    raise ValueError(
        "No records remain after preprocessing. "
        "Check the required fields and source data."
    )


# ================================================================
#  KEY BUSINESS METRICS
# ================================================================

total_records = len(analysis_df)

if "Applicant_ID" in analysis_df.columns:
    unique_applicants = analysis_df["Applicant_ID"].nunique()
else:
    unique_applicants = None

total_requested_amount = analysis_df["LoanAmount"].sum()

average_requested_amount = analysis_df["LoanAmount"].mean()

eligible_records = (
    analysis_df["Eligibility_Status"]
    .eq("Eligible")
    .sum()
)

overall_eligibility_rate = (
    eligible_records / total_records * 100
)

applications_by_product = (
    analysis_df["Loan_Type"]
    .value_counts()
    .sort_values(ascending=False)
)

eligibility_by_product = (
    analysis_df.assign(
        Is_Eligible=analysis_df["Eligibility_Status"].eq("Eligible")
    )
    .groupby("Loan_Type")
    .agg(
        Total_Applications=("Is_Eligible", "size"),
        Eligible_Applications=("Is_Eligible", "sum")
    )
)

eligibility_by_product["Eligibility_Rate"] = (
    eligibility_by_product["Eligible_Applications"]
    / eligibility_by_product["Total_Applications"]
    * 100
)

# Monthly applications: count records for each calendar month
monthly_applications = (
    analysis_df
    .set_index("Application_Date")
    .resample("MS")
    .size()
)

# Fill missing months between the first and last application date
if not monthly_applications.empty:
    full_month_range = pd.date_range(
        start=monthly_applications.index.min(),
        end=monthly_applications.index.max(),
        freq="MS"
    )

    monthly_applications = monthly_applications.reindex(
        full_month_range,
        fill_value=0
    )

    monthly_applications.index.name = "Application_Month"

# Peak month, based on the retained data
if not monthly_applications.empty:
    peak_month = monthly_applications.idxmax()
    peak_month_count = int(monthly_applications.max())
else:
    peak_month = None
    peak_month_count = 0

# Loan amount compared with policy maximum by product.
# These are sums of the provided fields, not individual policy decisions.
loan_policy_comparison = (
    analysis_df
    .groupby("Loan_Type")
    .agg(
        Total_Requested_Loan=("LoanAmount", "sum"),
        Total_Policy_Maximum=("Policy_Max_Loan", "sum"),
        Application_Count=("LoanAmount", "size")
    )
)

loan_policy_comparison["Average_Requested_Loan"] = (
    analysis_df.groupby("Loan_Type")["LoanAmount"].mean()
)

loan_policy_comparison["Average_Policy_Maximum"] = (
    analysis_df.groupby("Loan_Type")["Policy_Max_Loan"].mean()
)

# Save business metrics to CSV
summary_metrics = {
    "Metric": [
        "Total Records Analyzed",
        "Unique Applicants",
        "Total Requested Loan Amount",
        "Average Requested Loan Amount",
        "Eligible Records",
        "Overall Eligibility Rate",
        "Peak Application Month",
        "Applications in Peak Month"
    ],
    "Value": [
        total_records,
        unique_applicants if unique_applicants is not None else "Not available",
        round(total_requested_amount, 2),
        round(average_requested_amount, 2),
        int(eligible_records),
        round(overall_eligibility_rate, 2),
        peak_month.strftime("%B %Y") if peak_month is not None else "Not available",
        peak_month_count
    ]
}

summary_df = pd.DataFrame(summary_metrics)

summary_df.to_csv(
    OUTPUT_DIR / "business_summary_metrics.csv",
    index=False
)

applications_by_product.rename(
    "Application_Count"
).to_csv(
    OUTPUT_DIR / "applications_by_product.csv"
)

eligibility_by_product.to_csv(
    OUTPUT_DIR / "eligibility_by_product.csv"
)

loan_policy_comparison.to_csv(
    OUTPUT_DIR / "loan_policy_comparison.csv"
)

monthly_applications.rename(
    "Application_Count"
).to_csv(
    OUTPUT_DIR / "monthly_applications.csv"
)

# Print actual results
print("\n" + "=" * 65)
print("BUSINESS SUMMARY METRICS")
print("=" * 65)

print(f"Total records analyzed       : {total_records:,}")

if unique_applicants is not None:
    print(f"Unique applicants            : {unique_applicants:,}")

print(f"Total requested loan amount  : ₹{total_requested_amount:,.2f}")
print(f"Average requested loan       : ₹{average_requested_amount:,.2f}")
print(f"Eligible records             : {eligible_records:,}")
print(f"Overall eligibility rate     : {overall_eligibility_rate:.2f}%")

if peak_month is not None:
    print(
        f"Peak application month       : "
        f"{peak_month.strftime('%B %Y')}"
    )
    print(f"Applications in peak month   : {peak_month_count:,}")

print("\nApplications by loan product:")
print(applications_by_product.to_string())

print("\nEligibility rate by loan product:")
print(
    eligibility_by_product[
        ["Total_Applications", "Eligible_Applications", "Eligibility_Rate"]
    ].round(2).to_string()
)

print("\nRequested loan versus policy maximum by product:")
print(loan_policy_comparison.round(2).to_string())


# ================================================================
# CHART 1 — CUSTOMER ELIGIBILITY DISTRIBUTION
# ================================================================

eligibility_counts = (
    analysis_df["Eligibility_Status"]
    .value_counts()
)

fig, ax = plt.subplots()

eligibility_counts.plot(
    kind="bar",
    ax=ax
)

ax.set_title("Customer Eligibility Distribution")
ax.set_xlabel("Eligibility Status")
ax.set_ylabel("Number of Records")
ax.tick_params(axis="x", rotation=20)
ax.grid(axis="y", alpha=0.25)

for container in ax.containers:
    ax.bar_label(container, fmt="%.0f", padding=3)

fig.tight_layout()

fig.savefig(
    OUTPUT_DIR / "chart_1_eligibility_distribution.png",
    bbox_inches="tight"
)


# ================================================================
# CHART 2 — APPLICATIONS BY LOAN PRODUCT
# ================================================================

fig, ax = plt.subplots()

applications_by_product.plot(
    kind="bar",
    ax=ax
)

ax.set_title("Loan Applications by Product")
ax.set_xlabel("Loan Product")
ax.set_ylabel("Number of Applications")
ax.tick_params(axis="x", rotation=25)
ax.grid(axis="y", alpha=0.25)

for container in ax.containers:
    ax.bar_label(container, fmt="%.0f", padding=3)

fig.tight_layout()

fig.savefig(
    OUTPUT_DIR / "chart_2_applications_by_product.png",
    bbox_inches="tight"
)


# ================================================================
# CHART 3 — MONTHLY LOAN APPLICATION TIME SERIES
# ================================================================

fig, ax = plt.subplots()

ax.plot(
    monthly_applications.index,
    monthly_applications.values,
    marker="o",
    linewidth=2
)

ax.set_title("Monthly Loan Application Trend")
ax.set_xlabel("Application Month")
ax.set_ylabel("Number of Applications")

ax.xaxis.set_major_locator(
    mdates.AutoDateLocator()
)

ax.xaxis.set_major_formatter(
    mdates.ConciseDateFormatter(ax.xaxis.get_major_locator())
)

ax.grid(alpha=0.25)

# Annotate values when there are a manageable number of months
if len(monthly_applications) <= 24:
    for date, value in monthly_applications.items():
        ax.annotate(
            str(value),
            (date, value),
            textcoords="offset points",
            xytext=(0, 7),
            ha="center",
            fontsize=8
        )

fig.tight_layout()

fig.savefig(
    OUTPUT_DIR / "chart_3_monthly_time_series.png",
    bbox_inches="tight"
)


# ================================================================
# CHART 4 — ELIGIBILITY RATE BY LOAN PRODUCT
# ================================================================

product_rates = (
    eligibility_by_product["Eligibility_Rate"]
    .sort_values(ascending=False)
)

fig, ax = plt.subplots()

product_rates.plot(
    kind="bar",
    ax=ax
)

ax.set_title("Eligibility Rate by Loan Product")
ax.set_xlabel("Loan Product")
ax.set_ylabel("Eligibility Rate")
ax.set_ylim(0, 100)

ax.yaxis.set_major_formatter(
    mtick.PercentFormatter(xmax=100)
)

ax.tick_params(axis="x", rotation=25)
ax.grid(axis="y", alpha=0.25)

for container in ax.containers:
    ax.bar_label(
        container,
        fmt="%.1f%%",
        padding=3
    )

fig.tight_layout()

fig.savefig(
    OUTPUT_DIR / "chart_4_eligibility_rate.png",
    bbox_inches="tight"
)


# ================================================================
# CHART 5 — REQUESTED LOAN VS POLICY MAXIMUM
# ================================================================

plot_data = loan_policy_comparison[
    ["Average_Requested_Loan", "Average_Policy_Maximum"]
].sort_index()

fig, ax = plt.subplots()

plot_data.plot(
    kind="bar",
    ax=ax
)

ax.set_title("Average Requested Loan vs Average Policy Maximum")
ax.set_xlabel("Loan Product")
ax.set_ylabel("Amount (₹)")
ax.tick_params(axis="x", rotation=25)

ax.yaxis.set_major_formatter(
    mtick.FuncFormatter(
        lambda value, position: f"₹{value:,.0f}"
    )
)

ax.legend(
    ["Average Requested Loan", "Average Policy Maximum"]
)

ax.grid(axis="y", alpha=0.25)

fig.tight_layout()

fig.savefig(
    OUTPUT_DIR / "chart_5_requested_loan_vs_policy_maximum.png",
    bbox_inches="tight"
)


# ================================================================
# COMBINED DASHBOARD — ALL 5 CHARTS IN ONE IMAGE
# ================================================================

fig, axes = plt.subplots(
    2, 3,
    figsize=(18, 10)
)

fig.suptitle(
    "NOVATRUST BANK — EDA & TIME SERIES ANALYSIS",
    fontsize=18,
    fontweight="bold"
)

# Dashboard Chart 1
eligibility_counts.plot(
    kind="bar",
    ax=axes[0, 0]
)

axes[0, 0].set_title("1. Customer Eligibility")
axes[0, 0].set_xlabel("")
axes[0, 0].set_ylabel("Records")
axes[0, 0].tick_params(axis="x", rotation=20)
axes[0, 0].grid(axis="y", alpha=0.25)

# Dashboard Chart 2
applications_by_product.plot(
    kind="bar",
    ax=axes[0, 1]
)

axes[0, 1].set_title("2. Applications by Product")
axes[0, 1].set_xlabel("")
axes[0, 1].set_ylabel("Applications")
axes[0, 1].tick_params(axis="x", rotation=25)
axes[0, 1].grid(axis="y", alpha=0.25)

# Dashboard Chart 3 — Time series
axes[0, 2].plot(
    monthly_applications.index,
    monthly_applications.values,
    marker="o",
    linewidth=2
)

axes[0, 2].set_title("3. Monthly Application Trend")
axes[0, 2].set_xlabel("Month")
axes[0, 2].set_ylabel("Applications")
axes[0, 2].xaxis.set_major_formatter(
    mdates.ConciseDateFormatter(
        mdates.AutoDateLocator()
    )
)
axes[0, 2].grid(alpha=0.25)

# Dashboard Chart 4
product_rates.plot(
    kind="bar",
    ax=axes[1, 0]
)

axes[1, 0].set_title("4. Eligibility Rate by Product")
axes[1, 0].set_xlabel("")
axes[1, 0].set_ylabel("Eligibility Rate")
axes[1, 0].set_ylim(0, 100)
axes[1, 0].yaxis.set_major_formatter(
    mtick.PercentFormatter(xmax=100)
)
axes[1, 0].tick_params(axis="x", rotation=25)
axes[1, 0].grid(axis="y", alpha=0.25)

# Dashboard Chart 5
plot_data.plot(
    kind="bar",
    ax=axes[1, 1]
)

axes[1, 1].set_title("5. Requested Loan vs Policy Maximum")
axes[1, 1].set_xlabel("")
axes[1, 1].set_ylabel("Amount (₹)")
axes[1, 1].yaxis.set_major_formatter(
    mtick.FuncFormatter(
        lambda value, position: f"₹{value:,.0f}"
    )
)
axes[1, 1].tick_params(axis="x", rotation=25)
axes[1, 1].legend(
    ["Avg. Requested Loan", "Avg. Policy Maximum"],
    fontsize=8
)
axes[1, 1].grid(axis="y", alpha=0.25)

# Sixth panel is text, not an additional chart
axes[1, 2].axis("off")

axes[1, 2].text(
    0.02, 0.90,
    "KEY BUSINESS METRICS",
    fontsize=13,
    fontweight="bold"
)

insights = [
    f"Records analyzed: {total_records:,}",
    (
        f"Unique applicants: {unique_applicants:,}"
        if unique_applicants is not None
        else "Unique applicants: Not available"
    ),
    f"Eligible records: {eligible_records:,}",
    f"Eligibility rate: {overall_eligibility_rate:.1f}%",
    f"Total requested: ₹{total_requested_amount:,.0f}",
    f"Average requested: ₹{average_requested_amount:,.0f}",
    (
        f"Peak month: {peak_month.strftime('%b %Y')}"
        if peak_month is not None
        else "Peak month: Not available"
    ),
    f"Peak-month records: {peak_month_count:,}"
]

for index, line in enumerate(insights):
    axes[1, 2].text(
        0.02,
        0.78 - index * 0.09,
        line,
        fontsize=9
    )

fig.tight_layout(rect=[0, 0, 1, 0.94])

fig.savefig(
    OUTPUT_DIR / "figure_34_python_eda_dashboard.png",
    dpi=300,
    bbox_inches="tight"
)

# ================================================================
# FINISH
# ================================================================

# Close individual chart figures after saving.
# The combined dashboard is displayed at the end.
plt.close("all")

# Reopen and display the saved dashboard
dashboard_image = plt.imread(
    OUTPUT_DIR / "figure_34_python_eda_dashboard.png"
)

fig, ax = plt.subplots(figsize=(18, 10))
ax.imshow(dashboard_image)
ax.axis("off")
plt.tight_layout()
plt.show()

print("\n" + "=" * 65)
print("PYTHON ANALYSIS COMPLETED")
print("=" * 65)
print(f"Output folder: {OUTPUT_DIR.resolve()}")
print("\nGenerated files:")
for output_file in sorted(OUTPUT_DIR.iterdir()):
    print(" -", output_file.name)

print("\nNote:")
print(
    "The analysis uses the retained dataset after preprocessing. "
    "Monthly trends describe the supplied Application_Date values; "
    "they do not establish causes or predict future applications."
)

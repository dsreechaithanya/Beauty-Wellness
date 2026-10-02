"""Generate a researched Week 1 beauty and wellness report as a Word document.

Run with Python 3.10+ and no third-party packages:
    python beauty_wellness_week1.py
    python beauty_wellness_week1.py transactions.csv --output submission.docx

The optional CSV must contain: customer_id, order_id, order_date, net_revenue,
channel, category, promo_flag. Dates use YYYY-MM-DD; promo_flag uses 1/0 or
true/false. If no CSV is given, a clearly labeled synthetic demonstration is used.
"""

from __future__ import annotations

import argparse
import html
import math
import random
import statistics
import sys
import zipfile
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET


TITLE = "Predicting First-to-Second Purchase in Skincare E-Commerce"
REPORT_DATE = "October 2026"
REQUIRED_COLUMNS = {
    "customer_id", "order_id", "order_date", "net_revenue", "channel",
    "category", "promo_flag",
}
NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XML_NS = "http://www.w3.org/XML/1998/namespace"
ET.register_namespace("w", NS)
ET.register_namespace("r", REL_NS)


def wtag(name: str) -> str:
    return "{" + NS + "}" + name


def make_run(text: str, *, bold: bool = False, italic: bool = False,
             font: str | None = None, size: int | None = None,
             color: str | None = None) -> ET.Element:
    run = ET.Element(wtag("r"))
    props = ET.SubElement(run, wtag("rPr"))
    if bold:
        ET.SubElement(props, wtag("b"))
    if italic:
        ET.SubElement(props, wtag("i"))
    if font:
        fonts = ET.SubElement(props, wtag("rFonts"))
        for key in ("ascii", "hAnsi", "cs"):
            fonts.set(wtag(key), font)
    if size:
        ET.SubElement(props, wtag("sz")).set(wtag("val"), str(size))
    if color:
        ET.SubElement(props, wtag("color")).set(wtag("val"), color)
    node = ET.SubElement(run, wtag("t"))
    if text.startswith(" ") or text.endswith(" "):
        node.set("{" + XML_NS + "}space", "preserve")
    node.text = text
    return run


def paragraph(text: str = "", style: str | None = None, *, bold: bool = False,
              italic: bool = False, font: str | None = None,
              size: int | None = None, color: str | None = None,
              keep_next: bool = False) -> ET.Element:
    para = ET.Element(wtag("p"))
    props = ET.SubElement(para, wtag("pPr"))
    if style:
        ET.SubElement(props, wtag("pStyle")).set(wtag("val"), style)
    spacing = ET.SubElement(props, wtag("spacing"))
    spacing.set(wtag("after"), "120")
    spacing.set(wtag("line"), "276")
    spacing.set(wtag("lineRule"), "auto")
    if keep_next:
        ET.SubElement(props, wtag("keepNext"))
    if text:
        para.append(make_run(text, bold=bold, italic=italic, font=font,
                             size=size, color=color))
    return para


def page_break() -> ET.Element:
    para = ET.Element(wtag("p"))
    run = ET.SubElement(para, wtag("r"))
    ET.SubElement(run, wtag("br")).set(wtag("type"), "page")
    return para


def make_table(headers: list[str], rows: list[list[str]]) -> ET.Element:
    table = ET.Element(wtag("tbl"))
    props = ET.SubElement(table, wtag("tblPr"))
    borders = ET.SubElement(props, wtag("tblBorders"))
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = ET.SubElement(borders, wtag(edge))
        border.set(wtag("val"), "single")
        border.set(wtag("sz"), "4")
        border.set(wtag("color"), "D6DEE5")
    widths = [int(9360 / len(headers))] * len(headers)
    grid = ET.SubElement(table, wtag("tblGrid"))
    for width in widths:
        ET.SubElement(grid, wtag("gridCol")).set(wtag("w"), str(width))
    for row_number, values in enumerate([headers, *rows]):
        row = ET.SubElement(table, wtag("tr"))
        if row_number == 0:
            row_props = ET.SubElement(row, wtag("trPr"))
            ET.SubElement(row_props, wtag("tblHeader"))
        for index, value in enumerate(values):
            cell = ET.SubElement(row, wtag("tc"))
            cell_props = ET.SubElement(cell, wtag("tcPr"))
            cell_width = ET.SubElement(cell_props, wtag("tcW"))
            cell_width.set(wtag("w"), str(widths[index]))
            cell_width.set(wtag("type"), "dxa")
            if row_number == 0:
                shade = ET.SubElement(cell_props, wtag("shd"))
                shade.set(wtag("fill"), "183B45")
                shade.set(wtag("val"), "clear")
            cell_para = paragraph(str(value), bold=row_number == 0,
                                  size=18 if row_number == 0 else 17,
                                  color="FFFFFF" if row_number == 0 else "24333A")
            cell.append(cell_para)
    return table


def build_report_blocks(analysis: dict[str, Any], source_code: str) -> list[ET.Element]:
    blocks: list[ET.Element] = []

    def add(text: str, style: str | None = None, **kwargs: Any) -> None:
        blocks.append(paragraph(text, style, **kwargs))

    def heading(text: str, level: int = 1) -> None:
        blocks.append(paragraph(text, f"Heading{level}", keep_next=True))

    def bullets(items: list[str]) -> None:
        for item in items:
            add("\u2022  " + item)

    add("BEAUTY & WELLNESS  /  WEEK 1", "Subtitle", color="2D6A65", bold=True)
    add(TITLE, "Title")
    add("Business problem definition, research questions, evidence, and analysis plan", "Subtitle")
    add(f"Prepared {REPORT_DATE}  |  Data science orientation project", "Subtitle", color="58717A")
    heading("Executive Summary")
    add(
        "This proposal addresses a focused growth problem for a direct-to-consumer skincare retailer: "
        "which first-time customers are less likely to place a second order within 90 days, and can a "
        "relevant, consent-based follow-up improve that outcome? Beauty is a large, digitally shifting "
        "and competitive category. The business opportunity is not simply to send more promotions; it is "
        "to identify where the first-purchase journey breaks down and test helpful interventions without "
        "mistaking correlation for causation."
    )
    add(
        "The proposed first stage is a customer-level, 90-day repeat-purchase analysis using order, "
        "product, channel, promotion, and (where permitted) marketing-exposure data. Exploratory analysis "
        "will establish cohort-level baselines; a transparent logistic-regression baseline will estimate "
        "repeat propensity; and a randomized holdout experiment will test whether an intervention creates "
        "incremental repeat orders and gross margin. The included Python file runs a standard-library-only "
        "prototype and generates this Word report. Its default data are synthetic and are not evidence "
        "about real consumers."
    )
    heading("1. Industry Context and Relevance")
    add(
        "Beauty and wellness are increasingly discussed together: consumers may seek appearance, "
        "self-expression, routine, and a sense of well-being from products and services. This project "
        "deliberately narrows the scope to skincare retail transactions, where repeat orders, product "
        "categories, price, promotions, and digital touchpoints can be measured. It does not infer "
        "medical outcomes or treat purchase behavior as a measure of health."
    )
    add(
        "A 2023 McKinsey / Business of Fashion industry report estimated the global beauty market at "
        "about $430 billion in 2022 and projected about $580 billion by 2027. Those are dated estimates "
        "and forecasts, not realized 2026 sales. The same report described beauty e-commerce as having "
        "nearly quadrupled from 2015 to 2022 and exceeding 20% of sales in 2022; it also reported that "
        "42% of surveyed consumers across six named markets said they enjoyed trying new brands. These "
        "signals make customer acquisition, product discovery, and retention practical marketing questions, "
        "while also underlining how quickly category averages can become stale [1]."
    )
    add(
        "The World Health Organization describes mental well-being as shaped by interacting individual, "
        "social, and environmental factors [6]. A retail purchase record cannot measure a person's "
        "well-being, skin health, or treatment need; this proposal therefore uses wellness as market "
        "context, not as a customer-level label or outcome."
    )
    add(
        "Customer journeys span multiple touchpoints and channels; Lemon and Verhoef argue that firms "
        "need to understand experience across that journey rather than optimize isolated interactions [2]. "
        "Wedel and Kannan review how structured and unstructured data can support marketing decisions, "
        "including personalization, while highlighting privacy and data-security concerns [3]. For this "
        "reason, the proposed unit of analysis is the customer's first-purchase cohort, and the business "
        "outcome is evaluated alongside customer experience and contact-related guardrails."
    )
    heading("2. Business Problem and Scope")
    heading("Problem statement", 2)
    add(
        "A skincare e-commerce retailer has limited budget and attention for post-purchase communication. "
        "It does not know which observable first-order conditions are associated with a customer's next "
        "purchase, whether those patterns differ by acquisition channel, product category, promotion use, "
        "or customer cohort, or which follow-up action creates incremental value. Without a reliable "
        "baseline, teams risk over-discounting customers who would have returned anyway, ignoring customers "
        "who need product-use guidance, and confusing delayed replenishment with permanent churn."
    )
    heading("Scope and operational definition", 2)
    bullets([
        "Population: identifiable, eligible first-time purchasers of skincare from one retailer; exclude test orders, employees, and fraud according to documented business rules.",
        "Index event: each customer's earliest eligible completed order in the observation window.",
        "Primary outcome: at least one subsequent completed order from day 1 through day 90 after the index order.",
        "Observation rule: include a customer in the primary outcome denominator only when 90 full days of follow-up are observable; otherwise right-censor and exclude from the 90-day label.",
        "Unit: one customer, not one line item. Aggregate order lines to order level before defining the index order.",
        "Interpretation: a purchase propensity score ranks customers by observed likelihood; it does not establish why they return or prove that an offer will change their behavior."
    ])
    heading("3. Objectives and Success Criteria")
    bullets([
        "Describe first-order customers, order economics, and 90-day repeat rates by acquisition cohort, category, channel, and promotion status.",
        "Identify data-quality gaps, censoring, seasonality, returns, identity-resolution issues, and possible label leakage before modeling.",
        "Build an interpretable baseline model using only information available at the first order; compare it with a prevalence-only benchmark.",
        "Translate model scores into a feasible audience and customer-respectful action, then estimate incremental effects with randomized assignment.",
        "Judge the intervention on incremental 90-day repeat rate and contribution margin, with unsubscribe, complaint, return, and discount-cost guardrails."
    ])
    add(
        "Planning target (to be finalized with the business): identify a prospective test population large "
        "enough to detect a commercially meaningful improvement at 80% power and a 5% two-sided Type I "
        "error rate. Define the minimum detectable effect, baseline rate, allocation, and sample-size "
        "calculation before launch. Do not declare success from a model AUC alone."
    )
    heading("4. Research Questions and Hypotheses")
    blocks.append(make_table(
        ["Research question", "Testable hypothesis / decision"],
        [
            ["RQ1. What is the mature-cohort 90-day repeat-purchase rate, and how does it vary over time?", "H1. Repeat rates differ across first-purchase cohorts; quantify uncertainty and check calendar and product-mix changes before attributing the difference."],
            ["RQ2. Are first-order channel, category, basket value, and promotion associated with repeat purchase?", "H2. These first-order features provide predictive information beyond the overall repeat-rate baseline; assess out of time, not on training fit."],
            ["RQ3. Does a useful, non-discount follow-up outperform business as usual?", "H3. Among eligible first-time purchasers, randomized routine guidance or a replenishment reminder increases 90-day repeat rate versus a no-contact/business-as-usual holdout."],
            ["RQ4. Does a targeted incentive add enough value to justify its cost?", "H4. Any incremental repeat revenue must exceed discount, fulfillment, and contact costs; estimate treatment effects by pre-registered segments only."],
        ],
    ))
    add(
        "These are hypotheses, not assumed results. In particular, observed associations with promotion "
        "are confounded by targeting and customer intent; only randomized assignment supports a causal "
        "estimate of the intervention under the experiment's conditions."
    )
    heading("5. Data Plan and Variable Definitions")
    add(
        "The primary analytical table is one row per eligible order, joined to a pseudonymous customer "
        "key. The included CSV prototype requires the first seven fields below. A production study should "
        "retain a data dictionary, source system, refresh date, timezone, currency, and transformation "
        "logic for every field."
    )
    blocks.append(make_table(
        ["Variable", "Definition / role", "Source and cautions"],
        [
            ["customer_id", "Pseudonymous stable customer key; aggregation and split unit", "Commerce / identity graph; audit guest checkout and household collisions."],
            ["order_id, order_date", "Unique completed order and local/UTC timestamp; sort and deduplicate", "Commerce platform; standardize timezone, canceled orders, and date boundaries."],
            ["net_revenue", "Net merchandise revenue after discounts and refunds, excluding tax/shipping (document chosen convention)", "Order ledger; retain currency and refund adjustments; never mix gross and net definitions."],
            ["channel", "First-order acquisition or purchase channel", "Analytics / CRM; separate acquisition source from checkout channel where possible."],
            ["category / SKU", "Skincare product family and item-level product mix", "Product catalog; map discontinued, bundle, and reformulated SKUs."],
            ["promo_flag / discount", "Whether and how much a first-order incentive was applied", "Order pricing; promotion assignment is not random in historical data."],
            ["returns / stock status", "Refund, return, and stock-out indicators", "Order and inventory systems; may censor or prevent repurchase."],
            ["campaign exposure / consent", "Eligible, delivered, clicked, opted out, and consent status with timestamps", "CRM / messaging logs; consent and suppression rules take priority."],
            ["rating / review text (optional)", "Customer-reported rating or public review, only if lawfully collected", "Review platform / survey; selection bias, moderation, and sensitive-text risks."],
            ["repeat_90d (target)", "1 if a second completed order occurs in days 1-90; else 0 for mature cohorts", "Derived from order history after index date; never include future events as predictors."],
        ],
    ))
    add(
        "CSV input for the included script: customer_id, order_id, order_date (YYYY-MM-DD), net_revenue, "
        "channel, category, promo_flag. Additional fields require an intentional extension to the analysis; "
        "the starter code will not silently infer their meaning. Before running it, pre-filter to eligible "
        "completed orders, aggregate line items to one row per order, use one documented currency and "
        "revenue convention, and exclude test/fraud orders. The prototype cannot infer cancellations, "
        "returns, or currency from these seven required fields."
    )
    heading("6. Step-by-Step Analysis Approach")
    steps = [
        ("1. Align the decision", "Interview marketing, e-commerce, merchandising, finance, privacy, and customer-care stakeholders. Specify eligible population, intervention options, business-as-usual journeys, margin definition, and legal basis for data use."),
        ("2. Acquire and document data", "Extract order, item, customer, campaign, returns, and inventory records for a sufficiently long period to observe the 90-day outcome. Record schemas, source timestamps, consent status, and known platform changes."),
        ("3. Validate and prepare", "Check required fields, key uniqueness, date parsing, impossible or negative revenue, missingness, duplicate orders, currency, order status, identity joins, refunds, and right-censoring. Aggregate line items and apply exclusions consistently."),
        ("4. Explore before modeling", "Plot monthly mature-cohort repeat rates and confidence intervals; examine distributions, missingness, first-order value, discounts, category, channel, and repeat timing. Compare cohorts and segments with denominators and uncertainty, not just raw percentages."),
        ("5. Create leakage-safe features", "At the first-order cutoff, derive only available predictors: log first basket value, channel, category, discount flag, acquisition source, and customer-approved engagement signals recorded before the cutoff. Exclude future orders, post-purchase reviews collected later, campaign response after cutoff, and variables that encode the target."),
        ("6. Establish simple benchmarks", "Report the overall prevalence baseline and a regularized logistic-regression model. Use chronological train/validation/test cohorts, customer-level separation, class-aware metrics, calibration, and subgroup stability. Tune thresholds against contact capacity and expected unit economics."),
        ("7. Translate scores into a decision", "For a top-scored audience, compare expected incremental contribution margin with contact cost, discount cost, and operational constraints. Prediction ranks probability; it does not identify treatment responsiveness. Do not target solely by propensity if the goal is incremental lift."),
        ("8. Test an intervention", "Pre-register a randomized controlled experiment with a persistent customer-level holdout. Compare a relevant usage/routine message and/or replenishment reminder with business as usual; test discounts only if margin supports them. Stratify randomization on a small set of pre-treatment factors."),
        ("9. Evaluate and operationalize", "Report intention-to-treat lift, confidence intervals, incremental orders, contribution margin, and guardrails. Monitor drift and deliverability; retrain only on a documented schedule. Keep a human owner, rollback path, and model card."),
    ]
    for label, detail in steps:
        add(label, "Heading2", keep_next=True)
        add(detail)
    heading("7. Exploratory Analysis and Measurement")
    bullets([
        "Cohort table: index month, number of mature first-time customers, second purchasers, 90-day repeat rate, and a binomial confidence interval.",
        "Order profile: unique customers, orders, net revenue, average first-order value, discount distribution, category/channel mix, and return/refund rate.",
        "Repeat timing: cumulative share of customers with a second order by day 30, 60, 90, and 120; investigate category-specific replenishment cycles.",
        "Segment comparisons: report denominators and confidence intervals; adjust or qualify comparisons when channel, promotion, or product mix shifts over time.",
        "Model diagnostics: out-of-time ROC-AUC and precision/recall at operational capacity, Brier score or calibration plot, lift by score decile, and comparison against a majority/prevalence baseline.",
        "Experiment outcomes: 90-day repeat rate, incremental contribution margin per randomized customer, discount spend, returns, unsubscribes, complaints, and delivery failures."
    ])
    heading("8. Methods, Tools, and Why")
    blocks.append(make_table(
        ["Stage", "Methods / tools", "Reason and boundary"],
        [
            ["Prototype in this file", "Python 3.10+ standard library: csv, datetime, statistics, collections, zipfile, XML", "Runs without package installation; validates, summarizes, and creates a Word file. The optional small logistic baseline is transparent but not a production model."],
            ["Analyst-scale EDA", "pandas, NumPy, matplotlib / seaborn", "Efficient typed tabular cleaning, grouped summaries, reproducible plots, and missing-data checks."],
            ["Predictive baseline", "scikit-learn logistic regression; calibration and time-based validation", "Interpretable benchmark and established metrics; regularization, leakage controls, and calibration still require explicit design."],
            ["Text, if added", "Reviewed lexicon/topic or supervised sentiment analysis", "Useful for themes in reviews, but sarcasm, product-specific language, review selection, and health-claim interpretation make human validation essential."],
            ["Causal decision", "Randomized A/B test; power calculation; intention-to-treat estimate", "Tests whether a communication causes lift; observational model scores alone cannot answer this."],
        ],
    ))
    heading("9. Ethics, Privacy, and Risks")
    bullets([
        "Use the minimum necessary pseudonymous data; restrict access, retention, and exports. Confirm consent, lawful basis, platform terms, and applicable privacy rules before joining customer-level data.",
        "Do not infer diagnoses, skin conditions, age, ethnicity, or other sensitive traits from purchase or review text. Avoid sensitive or protected attributes in targeting; assess proxy and disparate-impact risks with the privacy/legal team.",
        "Beauty and wellness marketing can imply health benefits. The FTC states that health-related advertising claims must be truthful, not misleading, and appropriately substantiated; a purchase/review prediction is not clinical evidence [4]. Keep messages factual and do not automate medical advice.",
        "Historical promotion effects are confounded; recommendations based on these data are associative until experimentally tested. Track inventory, seasonality, price changes, and marketing-policy shifts.",
        "A 90-day non-repeat is not necessarily churn: consumption and replenishment vary by product, household sharing, product size, and channel. Compare alternative 60/120/180-day windows and use survival analysis in a production study.",
        "Review text and ratings are self-selected and can over-represent unusually positive or negative experiences. Do not treat sentiment as representative market opinion or as evidence that a cosmetic product produces a medical outcome."
    ])
    heading("10. Proposed 35-Hour Work Plan")
    blocks.append(make_table(
        ["Workstream", "Hours", "Deliverable / gate"],
        [
            ["Industry and academic source review", "5", "Annotated evidence table; date-stamped facts and citation list."],
            ["Stakeholder framing and metric definition", "4", "Decision brief, population, index event, outcome, and test guardrails."],
            ["Data inventory and quality plan", "5", "Data dictionary, access/consent checklist, and cohort eligibility rules."],
            ["EDA specification and prototype", "6", "Cohort, channel, category, promotion, and repeat-timing analysis."],
            ["Model and validation design", "5", "Baseline, leakage-safe feature list, temporal split, and evaluation plan."],
            ["Experiment and economics", "5", "Test arms, sample-size inputs, unit economics, and decision thresholds."],
            ["Synthesis, review, and final document", "5", "Stakeholder review, limitations, references, and polished submission."],
            ["Total", "35", "Conceptual plan; actual time depends on data access and course requirements."],
        ],
    ))
    heading("11. Expected Deliverables and Decision")
    bullets([
        "A validated customer/order data dictionary and cohort eligibility specification.",
        "An exploratory report with mature-cohort repeat rates, segment comparisons, confidence intervals, and repeat-timing curves.",
        "A leakage-safe out-of-time baseline model and a documented assessment of whether the model adds value over simple rules.",
        "A pre-registered experiment plan with sample-size assumptions, holdout design, economics, and customer-experience guardrails.",
        "A launch / revise / stop recommendation based on incremental contribution margin and customer outcomes, not on accuracy metrics alone."
    ])
    add(
        "Recommendation: proceed first with data validation and descriptive cohort analysis. Build the "
        "predictive baseline only if identifiers, mature follow-up, and pre-purchase features are reliable. "
        "Use random assignment to establish intervention value before scaling personalization."
    )
    heading("Conclusion")
    add(
        "First-to-second purchase is a bounded, business-relevant entry point into beauty and wellness "
        "analytics. It connects customer journey research to a concrete marketing decision, supports a "
        "reproducible exploratory workflow, and creates a natural bridge from prediction to causal testing. "
        "The central discipline is to distinguish market context from customer-level evidence, prediction "
        "from explanation, and association from incremental impact."
    )
    heading("References")
    references = [
        "[1] Berg, A., Hudson, S., Weaver, K. K., Pacchia, M. L., & Amed, I. (2023, May 22). The beauty market in 2023: A special State of Fashion report. McKinsey & Company with The Business of Fashion. https://www.mckinsey.com/industries/consumer-packaged-goods/our-insights/the-beauty-market-in-2023-a-special-state-of-fashion-report",
        "[2] Lemon, K. N., & Verhoef, P. C. (2016). Understanding customer experience throughout the customer journey. Journal of Marketing, 80(6), 69-96. https://doi.org/10.1509/jm.15.0420",
        "[3] Wedel, M., & Kannan, P. K. (2016). Marketing analytics for data-rich environments. Journal of Marketing, 80(6), 97-121. https://doi.org/10.1509/jm.15.0413",
        "[4] Federal Trade Commission. (2022, December). Health products compliance guidance. https://www.ftc.gov/business-guidance/resources/health-products-compliance-guidance",
        "[5] Neslin, S. A., Gupta, S., Kamakura, W., Lu, J., & Mason, C. H. (2006). Defection detection: Measuring and understanding the predictive accuracy of customer churn models. Journal of Marketing Research, 43(2), 204-211. https://doi.org/10.1509/jmkr.43.2.204",
        "[6] World Health Organization. (2022). World mental health report: Transforming mental health for all. https://www.who.int/publications/i/item/9789240049338",
    ]
    for reference in references:
        add(reference, size=18)
    add(
        "Source note: Industry figures above are reproduced as dated estimates/forecasts from the cited "
        "2023 report. Peer-reviewed article metadata were checked against Crossref DOI records. Sources "
        "provide conceptual context; they do not validate the proposed hypotheses or predict this "
        "retailer's performance. Accessed October 2026.", italic=True, size=18, color="58717A"
    )
    blocks.append(page_break())
    heading("Appendix A. Analysis Output")
    add(analysis["label"], italic=True, color="58717A")
    for line in analysis["summary"]:
        add(line, font="Consolas", size=18)
    for line in analysis.get("model_summary", []):
        add(line, font="Consolas", size=18)
    heading("Appendix B. Run Instructions")
    add("1. Install Python 3.10 or newer. No third-party packages are required.")
    add("2. In VS Code, open this .py file and select Run Python File, or run: python beauty_wellness_week1.py")
    add("3. The script writes beauty_wellness_week1_report.docx beside itself. Open that .docx in Microsoft Word and submit it as the document deliverable.")
    add("4. To analyze your own file: python beauty_wellness_week1.py transactions.csv --output final_report.docx. Use the exact seven required CSV columns listed at the start of this report.")
    add("5. The report's final appendix contains the complete Python source. Keep the .py as your runnable single-file source; submit the generated .docx if the assignment requires a Word document.")
    blocks.append(page_break())
    heading("Appendix C. Complete Single-File Python Source")
    add("The executable source used to generate this report follows. The default fixture is synthetic; replace it with an appropriately authorized, documented CSV before drawing business conclusions.", italic=True)
    for line in source_code.splitlines():
        code_para = paragraph(line if line else " ", font="Consolas", size=14)
        p_props = code_para.find(wtag("pPr"))
        if p_props is not None:
            spacing = p_props.find(wtag("spacing"))
            if spacing is not None:
                spacing.set(wtag("after"), "0")
                spacing.set(wtag("line"), "220")
        blocks.append(code_para)
    return blocks


def styles_xml() -> bytes:
    root = ET.Element(wtag("styles"))
    defaults = ET.SubElement(root, wtag("docDefaults"))
    run_defaults = ET.SubElement(defaults, wtag("rPrDefault"))
    run_props = ET.SubElement(run_defaults, wtag("rPr"))
    fonts = ET.SubElement(run_props, wtag("rFonts"))
    fonts.set(wtag("ascii"), "Georgia")
    fonts.set(wtag("hAnsi"), "Georgia")
    ET.SubElement(run_props, wtag("sz")).set(wtag("val"), "20")
    paragraph_defaults = ET.SubElement(defaults, wtag("pPrDefault"))
    ET.SubElement(paragraph_defaults, wtag("pPr"))
    definitions = [
        ("Normal", "Calibri", 20, "24333A", False),
        ("Title", "Georgia", 34, "183B45", True),
        ("Subtitle", "Calibri", 21, "42616A", False),
        ("Heading1", "Georgia", 27, "183B45", True),
        ("Heading2", "Calibri", 22, "2D6A65", True),
    ]
    for name, font, size, color, bold in definitions:
        style = ET.SubElement(root, wtag("style"))
        style.set(wtag("type"), "paragraph")
        style.set(wtag("styleId"), name)
        ET.SubElement(style, wtag("name")).set(wtag("val"), name)
        if name != "Normal":
            ET.SubElement(style, wtag("basedOn")).set(wtag("val"), "Normal")
        props = ET.SubElement(style, wtag("rPr"))
        style_fonts = ET.SubElement(props, wtag("rFonts"))
        style_fonts.set(wtag("ascii"), font)
        style_fonts.set(wtag("hAnsi"), font)
        ET.SubElement(props, wtag("sz")).set(wtag("val"), str(size))
        ET.SubElement(props, wtag("color")).set(wtag("val"), color)
        if bold:
            ET.SubElement(props, wtag("b"))
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def generate_docx(path: Path, blocks: list[ET.Element]) -> None:
    document = ET.Element(wtag("document"))
    body = ET.SubElement(document, wtag("body"))
    for block in blocks:
        body.append(block)
    section = ET.SubElement(body, wtag("sectPr"))
    page_size = ET.SubElement(section, wtag("pgSz"))
    page_size.set(wtag("w"), "12240")
    page_size.set(wtag("h"), "15840")
    margins = ET.SubElement(section, wtag("pgMar"))
    margins.set(wtag("top"), "1080")
    margins.set(wtag("right"), "1080")
    margins.set(wtag("bottom"), "1080")
    margins.set(wtag("left"), "1080")
    margins.set(wtag("header"), "540")
    margins.set(wtag("footer"), "540")
    margins.set(wtag("gutter"), "0")
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        '</Types>'
    )
    package_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
        '</Relationships>'
    )
    document_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '</Relationships>'
    )
    core = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        f'<dc:title>{html.escape(TITLE)}</dc:title>'
        '<dc:subject>Beauty and wellness data science problem definition</dc:subject>'
        '<dc:creator>Week 1 Data Science Orientation</dc:creator>'
        f'<dcterms:created xsi:type="dcterms:W3CDTF">{date.today().isoformat()}T00:00:00Z</dcterms:created>'
        '</cp:coreProperties>'
    )
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", package_rels)
        archive.writestr("word/_rels/document.xml.rels", document_rels)
        archive.writestr("word/document.xml", ET.tostring(document, encoding="utf-8", xml_declaration=True))
        archive.writestr("word/styles.xml", styles_xml())
        archive.writestr("docProps/core.xml", core)


def parse_bool(value: str) -> bool:
    normalized = value.strip().casefold()
    if normalized in {"1", "true", "yes", "y"}:
        return True
    if normalized in {"0", "false", "no", "n", ""}:
        return False
    raise ValueError(f"promo_flag must be 1/0 or true/false, not {value!r}")


def load_transactions(path: Path) -> tuple[list[dict[str, Any]], int]:
    import csv

    transactions: list[dict[str, Any]] = []
    skipped_duplicates = 0
    seen_order_ids: set[str] = set()
    with path.open("r", encoding="utf-8-sig", newline="") as stream:
        reader = csv.DictReader(stream)
        columns = set(reader.fieldnames or [])
        missing = sorted(REQUIRED_COLUMNS - columns)
        if missing:
            raise ValueError("CSV is missing required columns: " + ", ".join(missing))
        for line_number, row in enumerate(reader, start=2):
            try:
                customer_id = (row.get("customer_id") or "").strip()
                order_id = (row.get("order_id") or "").strip()
                if not customer_id or not order_id:
                    raise ValueError("customer_id and order_id cannot be blank")
                if order_id in seen_order_ids:
                    skipped_duplicates += 1
                    continue
                order_date = date.fromisoformat((row.get("order_date") or "").strip())
                revenue = float((row.get("net_revenue") or "").strip())
                if not math.isfinite(revenue):
                    raise ValueError("net_revenue must be a finite number")
                if revenue < 0:
                    raise ValueError("net_revenue cannot be negative in this prototype")
                channel = (row.get("channel") or "Unknown").strip() or "Unknown"
                category = (row.get("category") or "Unknown").strip() or "Unknown"
                promo = parse_bool(row.get("promo_flag") or "")
            except (ValueError, TypeError) as error:
                raise ValueError(f"CSV row {line_number}: {error}") from error
            seen_order_ids.add(order_id)
            transactions.append({
                "customer_id": customer_id,
                "order_id": order_id,
                "order_date": order_date,
                "net_revenue": revenue,
                "channel": channel,
                "category": category,
                "promo_flag": promo,
            })
    if not transactions:
        raise ValueError("CSV contains no valid transactions")
    return transactions, skipped_duplicates


def synthetic_transactions() -> list[dict[str, Any]]:
    rng = random.Random(20261002)
    transactions: list[dict[str, Any]] = []
    channels = ["Online", "Social", "Referral"]
    categories = ["Cleanser", "Moisturizer", "Serum", "Sunscreen"]
    start = date(2024, 1, 1)
    for index in range(72):
        customer = f"demo_{index + 1:03d}"
        first_date = start + timedelta(days=index * 4)
        promo = index % 3 == 0
        channel = channels[index % len(channels)]
        category = categories[(index * 3) % len(categories)]
        first_value = round(rng.uniform(18, 92), 2)
        transactions.append({
            "customer_id": customer,
            "order_id": f"demo_order_{index + 1:03d}_1",
            "order_date": first_date,
            "net_revenue": first_value,
            "channel": channel,
            "category": category,
            "promo_flag": promo,
        })
        will_repeat = ((index * 7) % 11) < 6
        if will_repeat:
            repeat_day = 18 + ((index * 13) % 110)
            transactions.append({
                "customer_id": customer,
                "order_id": f"demo_order_{index + 1:03d}_2",
                "order_date": first_date + timedelta(days=repeat_day),
                "net_revenue": round(rng.uniform(15, 78), 2),
                "channel": channel,
                "category": category,
                "promo_flag": False,
            })
    return transactions


def auc_score(labels: list[int], probabilities: list[float]) -> float | None:
    positives = [score for label, score in zip(labels, probabilities) if label == 1]
    negatives = [score for label, score in zip(labels, probabilities) if label == 0]
    if not positives or not negatives:
        return None
    wins = sum(1.0 if positive > negative else 0.5 if positive == negative else 0.0
               for positive in positives for negative in negatives)
    return wins / (len(positives) * len(negatives))


def fit_logistic(train_x: list[list[float]], train_y: list[int]) -> tuple[list[float], list[float], list[float]]:
    feature_count = len(train_x[0])
    means = [statistics.fmean(row[index] for row in train_x) for index in range(feature_count)]
    scales = []
    for index, mean in enumerate(means):
        variance = statistics.fmean((row[index] - mean) ** 2 for row in train_x)
        scales.append(math.sqrt(variance) or 1.0)
    scaled = [[(value - means[index]) / scales[index] for index, value in enumerate(row)]
              for row in train_x]
    weights = [0.0] * (feature_count + 1)
    learning_rate = 0.12
    for _ in range(1600):
        gradients = [0.0] * len(weights)
        for row, label in zip(scaled, train_y):
            score = weights[0] + sum(weight * value for weight, value in zip(weights[1:], row))
            probability = 1.0 / (1.0 + math.exp(-max(-30.0, min(30.0, score))))
            error = probability - label
            gradients[0] += error
            for index, value in enumerate(row, start=1):
                gradients[index] += error * value
        weights = [weight - learning_rate * gradient / len(train_y)
                   for weight, gradient in zip(weights, gradients)]
    return weights, means, scales


def predict_logistic(rows: list[list[float]], weights: list[float],
                     means: list[float], scales: list[float]) -> list[float]:
    probabilities = []
    for row in rows:
        standardized = [(value - means[index]) / scales[index]
                        for index, value in enumerate(row)]
        score = weights[0] + sum(weight * value for weight, value in zip(weights[1:], standardized))
        probabilities.append(1.0 / (1.0 + math.exp(-max(-30.0, min(30.0, score)))))
    return probabilities


def analyze(transactions: list[dict[str, Any]], duplicates: int, label: str) -> dict[str, Any]:
    transactions = sorted(transactions, key=lambda row: (row["order_date"], row["order_id"]))
    by_customer: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for transaction in transactions:
        by_customer[transaction["customer_id"]].append(transaction)
    latest_date = max(row["order_date"] for row in transactions)
    earliest_date = min(row["order_date"] for row in transactions)
    first_orders = [orders[0] for orders in by_customer.values()]
    mature_orders = [order for order in first_orders
                     if order["order_date"] + timedelta(days=90) <= latest_date]
    repeat_labels = []
    for first in mature_orders:
        followup = [order for order in by_customer[first["customer_id"]][1:]
                    if first["order_date"] < order["order_date"] <= first["order_date"] + timedelta(days=90)]
        repeat_labels.append(int(bool(followup)))
    first_values = [order["net_revenue"] for order in first_orders]
    repeat_rate = statistics.fmean(repeat_labels) if repeat_labels else None
    lines = [
        f"Data source: {label}",
        f"Date range: {earliest_date.isoformat()} to {latest_date.isoformat()}",
        f"Orders: {len(transactions):,} | Unique customers: {len(by_customer):,} | Duplicate order IDs skipped: {duplicates:,}",
        f"First-order net revenue: ${sum(first_values):,.2f} | Mean first-order value: ${statistics.fmean(first_values):,.2f}",
        f"Mature first-purchase customers (90-day follow-up): {len(mature_orders):,} of {len(first_orders):,}",
        f"Observed 90-day repeat rate: {repeat_rate:.1%}" if repeat_rate is not None else "Observed 90-day repeat rate: not estimable (no mature customers)",
        "Interpretation: descriptive only; synthetic fixture or observational transaction data do not establish causality.",
    ]
    category_counts: Counter[str] = Counter()
    category_repeats: Counter[str] = Counter()
    channel_counts: Counter[str] = Counter()
    channel_repeats: Counter[str] = Counter()
    for first, label_value in zip(mature_orders, repeat_labels):
        category_counts[first["category"]] += 1
        category_repeats[first["category"]] += label_value
        channel_counts[first["channel"]] += 1
        channel_repeats[first["channel"]] += label_value
    lines.append("Mature cohort by first-order category (customers / repeat rate):")
    if category_counts:
        for category, count in sorted(category_counts.items()):
            lines.append(f"  {category}: {count} / {category_repeats[category] / count:.1%}")
    else:
        lines.append("  Not available")
    lines.append("Mature cohort by first-order channel (customers / repeat rate):")
    if channel_counts:
        for channel, count in sorted(channel_counts.items()):
            lines.append(f"  {channel}: {count} / {channel_repeats[channel] / count:.1%}")
    else:
        lines.append("  Not available")
    model_lines: list[str] = []
    if len(mature_orders) >= 30 and len(set(repeat_labels)) == 2:
        ordered = sorted(zip(mature_orders, repeat_labels), key=lambda pair: pair[0]["order_date"])
        split = max(10, min(len(ordered) - 10, int(len(ordered) * 0.7)))
        train, test = ordered[:split], ordered[split:]
        train_y = [value for _, value in train]
        test_y = [value for _, value in test]
        if len(set(train_y)) == 2 and len(set(test_y)) == 2:
            def features(order: dict[str, Any]) -> list[float]:
                return [math.log1p(order["net_revenue"]), float(order["promo_flag"]),
                        float(order["channel"].casefold() in {"online", "e-commerce", "ecommerce", "web"})]

            weights, means, scales = fit_logistic([features(order) for order, _ in train], train_y)
            probabilities = predict_logistic([features(order) for order, _ in test], weights, means, scales)
            majority_rate = statistics.fmean(train_y)
            baseline_predictions = [majority_rate] * len(test_y)
            auc = auc_score(test_y, probabilities)
            brier = statistics.fmean((prediction - actual) ** 2
                                     for prediction, actual in zip(probabilities, test_y))
            base_brier = statistics.fmean((prediction - actual) ** 2
                                          for prediction, actual in zip(baseline_predictions, test_y))
            model_lines = [
                "Illustrative logistic baseline: log(first-order value), promotion flag, and online-channel flag.",
                f"Chronological split: train={len(train):,}, test={len(test):,}; test ROC-AUC={auc:.3f}",
                f"Test Brier score={brier:.3f}; prevalence-only benchmark Brier score={base_brier:.3f}.",
                "Caution: this small prototype is not production-ready; metrics are unstable on small samples and need confidence intervals, calibration review, and out-of-time validation.",
            ]
    if not model_lines:
        model_lines = [
            "Illustrative logistic baseline not fit: need at least 30 mature customers and both repeat outcomes in chronological training and test partitions.",
            "This is a sample-size safeguard, not evidence that predictors have no value.",
        ]
    return {"label": label, "summary": lines, "model_summary": model_lines}


def main() -> int:
    parser = argparse.ArgumentParser(description="Analyze skincare retention data and generate a Word report.")
    parser.add_argument("csv_path", nargs="?", type=Path,
                        help="Optional CSV with the seven required transaction columns.")
    parser.add_argument("--output", type=Path, default=None,
                        help="Output .docx path (default: beauty_wellness_week1_report.docx beside this script).")
    args = parser.parse_args()
    output_path = args.output or Path(__file__).resolve().with_name("beauty_wellness_week1_report.docx")
    try:
        if args.csv_path:
            transactions, duplicates = load_transactions(args.csv_path)
            label = f"USER CSV: {args.csv_path.name} (observational; verify authorization and definitions)"
        else:
            transactions = synthetic_transactions()
            duplicates = 0
            label = "SYNTHETIC DEMONSTRATION ONLY - not real customers, market research, or business findings"
        analysis = analyze(transactions, duplicates, label)
        source_code = Path(__file__).resolve().read_text(encoding="utf-8")
        blocks = build_report_blocks(analysis, source_code)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        generate_docx(output_path, blocks)
    except (OSError, ValueError, zipfile.BadZipFile) as error:
        print(f"Error: {error}", file=sys.stderr)
        return 2
    print(f"Word report created: {output_path}")
    for line in analysis["summary"]:
        print(line)
    for line in analysis["model_summary"]:
        print(line)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
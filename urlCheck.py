import argparse, csv, datetime, html, io, json, os, pathlib, re, signal, sys, traceback, urllib.error, urllib.parse, urllib.request
from playwright.sync_api import sync_playwright

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


# --- Constants ---

bDefaultFullPage = True
bDefaultHeadless = False
bDefaultIgnoreHttpsErrors = True

iCdnTimeoutSec = 30
iCsvMaxHtmlLen = 32000
iDefaultNavTimeoutMs = 60000
iDefaultPostLoadDelayMs = 1500
iDefaultViewportHeight = 1440
iDefaultViewportWidth = 1600
iMaxErrorTextLen = 12000
iMaxTitleLen = 80
iNetworkIdleTimeoutMs = 8000
iSuffixDigits = 3

sAccessibilityInsightsUrl = "https://accessibilityinsights.io/docs/web/overview/"
sBrowserChannel = "msedge"
sCsvName = "report.csv"
sErrorReportName = "error.txt"
sFallbackTitle = "untitled-page"
sJsonName = "results.json"
sMsAccessibilityUrl = "https://learn.microsoft.com/accessibility/"
sProgramName = "urlCheck"
sProgramVersion = "1.7.0"
sReportName = "report.htm"
sReportWorkbookName = "report.xlsx"
sScreenshotName = "page.png"
sSourceName = "page.html"
sUsage = "Usage: urlCheck [options] <url, domain, local html file, or url-list text file>"
sUserAgent = "urlCheck/1.7.0 (+Playwright Python + axe-core)"
sWcagBaseUrl = "https://www.w3.org/WAI/WCAG22/Understanding/"


# --- Data tables ---

aAllowedLocalExtensions = [".htm", ".html", ".xhtml"]

aLocalFileExts = {
    ".asp", ".aspx", ".css", ".csv", ".docx", ".gif", ".htm", ".html",
    ".ico", ".jpeg", ".jpg", ".js", ".json", ".jsx", ".md", ".pdf",
    ".php", ".png", ".pptx", ".py", ".svg", ".ts", ".txt", ".xhtml",
    ".xlsm", ".xlsx", ".xml", ".zip",
}

reDomain = re.compile(
    r"^([a-zA-Z0-9]([a-zA-Z0-9\-]{0,61}[a-zA-Z0-9])?\.)+[a-zA-Z]{2,}(:[0-9]+)?(/[^\s]*)?$"
)
reIpAddress = re.compile(r"^(\d{1,3}\.){3}\d{1,3}(:\d+)?(/[^\s]*)?$")

aAxeCdnUrls = [
    "https://cdn.jsdelivr.net/npm/axe-core@4.11.0/axe.min.js",
    "https://unpkg.com/axe-core@4.11.0/axe.min.js",
]

aAxeRunOptions = {"resultTypes": ["violations", "incomplete"]}

aGlossaryRows = [
    ["axe-core", "Deque Systems accessibility testing engine that runs rules against the browser DOM."],
    ["impact", "Axe severity estimate for a rule or node result, such as critical or serious."],
    ["inapplicable", "A rule that did not apply to this page or document."],
    ["incomplete", "A result that needs manual review because automated certainty was not possible."],
    ["node", "A single DOM target reported by axe-core under a rule result."],
    ["outcome", "The top-level axe result bucket: violations, incomplete, passes, or inapplicable."],
    ["pass", "A rule that applied and passed automated checks."],
    ["rule", "An axe-core test identified by a rule ID, help text, description, tags, and node results."],
    ["target", "A CSS selector or selector path returned by axe-core for a specific node result."],
    ["violation", "A confirmed automated accessibility issue found by axe-core."],
    ["WCAG reference", "A success criterion reference derived from axe-core tags such as wcag143, rendered as 1.4.3 where possible."],
]

aProcedures = [
    "Open the supplied URL or local HTML file in the installed Microsoft Edge browser through the synchronous Playwright API.",
    "Wait for the page to finish loading, then pause briefly so late DOM updates are more likely to settle.",
    "Load axe-core from reliable public CDNs without requiring Node.js on the user system.",
    "Run axe-core with its default behavior and the minimal resultTypes option set used to return violations, incomplete, passes, and inapplicable results.",
    "Capture the page title, final URL, user agent, browser version, screenshot, page HTML, and structured accessibility findings.",
    "Write a static HTML report and a structured Excel workbook.",
]

aRelevantTagPrefixes = ["EN-301-549", "best-practice", "cat.", "section508", "TT", "wcag", "wcag2", "wcag21", "wcag22"]

aOutputSections = ["violations"]
aReportSections = ["violations", "incomplete", "passes", "inapplicable"]

dImpactEmoji = {"critical": "🔴", "serious": "🟠", "moderate": "🟡", "minor": "⚪"}

dImpactRank = {"critical": 1, "serious": 2, "moderate": 3, "minor": 4, "": 5, None: 5}

dOutcomeRank = {"violations": 1, "incomplete": 2, "passes": 3, "inapplicable": 4}

dSectionHeadings = {
    "violations": "Violations",
    "incomplete": "Needs Review",
    "passes": "Passes",
    "inapplicable": "Inapplicable",
}

# WCAG 2.0 / 2.1 / 2.2 success criteria: SC number -> (short name, level, principle)
dWcagSc = {
    "1.1.1": ("Non-text Content",                         "A",   "Perceivable"),
    "1.2.1": ("Audio-only and Video-only (Prerecorded)",  "A",   "Perceivable"),
    "1.2.2": ("Captions (Prerecorded)",                   "A",   "Perceivable"),
    "1.2.3": ("Audio Description or Media Alternative",   "A",   "Perceivable"),
    "1.2.4": ("Captions (Live)",                          "AA",  "Perceivable"),
    "1.2.5": ("Audio Description (Prerecorded)",          "AA",  "Perceivable"),
    "1.2.6": ("Sign Language (Prerecorded)",              "AAA", "Perceivable"),
    "1.2.7": ("Extended Audio Description (Prerecorded)", "AAA", "Perceivable"),
    "1.2.8": ("Media Alternative (Prerecorded)",          "AAA", "Perceivable"),
    "1.2.9": ("Audio-only (Live)",                        "AAA", "Perceivable"),
    "1.3.1": ("Info and Relationships",                   "A",   "Perceivable"),
    "1.3.2": ("Meaningful Sequence",                      "A",   "Perceivable"),
    "1.3.3": ("Sensory Characteristics",                  "A",   "Perceivable"),
    "1.3.4": ("Orientation",                              "AA",  "Perceivable"),
    "1.3.5": ("Identify Input Purpose",                   "AA",  "Perceivable"),
    "1.3.6": ("Identify Purpose",                         "AAA", "Perceivable"),
    "1.4.1": ("Use of Color",                             "A",   "Perceivable"),
    "1.4.2": ("Audio Control",                            "A",   "Perceivable"),
    "1.4.3": ("Contrast (Minimum)",                       "AA",  "Perceivable"),
    "1.4.4": ("Resize Text",                              "AA",  "Perceivable"),
    "1.4.5": ("Images of Text",                           "AA",  "Perceivable"),
    "1.4.6": ("Contrast (Enhanced)",                      "AAA", "Perceivable"),
    "1.4.7": ("Low or No Background Audio",               "AAA", "Perceivable"),
    "1.4.8": ("Visual Presentation",                      "AAA", "Perceivable"),
    "1.4.9": ("Images of Text (No Exception)",            "AAA", "Perceivable"),
    "1.4.10": ("Reflow",                                  "AA",  "Perceivable"),
    "1.4.11": ("Non-text Contrast",                       "AA",  "Perceivable"),
    "1.4.12": ("Text Spacing",                            "AA",  "Perceivable"),
    "1.4.13": ("Content on Hover or Focus",               "AA",  "Perceivable"),
    "2.1.1": ("Keyboard",                                 "A",   "Operable"),
    "2.1.2": ("No Keyboard Trap",                         "A",   "Operable"),
    "2.1.3": ("Keyboard (No Exception)",                  "AAA", "Operable"),
    "2.1.4": ("Character Key Shortcuts",                  "A",   "Operable"),
    "2.2.1": ("Timing Adjustable",                        "A",   "Operable"),
    "2.2.2": ("Pause, Stop, Hide",                        "A",   "Operable"),
    "2.2.3": ("No Timing",                                "AAA", "Operable"),
    "2.2.4": ("Interruptions",                            "AAA", "Operable"),
    "2.2.5": ("Re-authenticating",                        "AAA", "Operable"),
    "2.2.6": ("Timeouts",                                 "AAA", "Operable"),
    "2.3.1": ("Three Flashes or Below Threshold",         "A",   "Operable"),
    "2.3.2": ("Three Flashes",                            "AAA", "Operable"),
    "2.3.3": ("Animation from Interactions",              "AAA", "Operable"),
    "2.4.1": ("Bypass Blocks",                            "A",   "Operable"),
    "2.4.2": ("Page Titled",                              "A",   "Operable"),
    "2.4.3": ("Focus Order",                              "A",   "Operable"),
    "2.4.4": ("Link Purpose (In Context)",                "A",   "Operable"),
    "2.4.5": ("Multiple Ways",                            "AA",  "Operable"),
    "2.4.6": ("Headings and Labels",                      "AA",  "Operable"),
    "2.4.7": ("Focus Visible",                            "AA",  "Operable"),
    "2.4.8": ("Location",                                 "AAA", "Operable"),
    "2.4.9": ("Link Purpose (Link Only)",                 "AAA", "Operable"),
    "2.4.10": ("Section Headings",                        "AAA", "Operable"),
    "2.4.11": ("Focus Not Obscured (Minimum)",            "AA",  "Operable"),
    "2.4.12": ("Focus Not Obscured (Enhanced)",           "AAA", "Operable"),
    "2.4.13": ("Focus Appearance",                        "AA",  "Operable"),
    "2.5.1": ("Pointer Gestures",                         "A",   "Operable"),
    "2.5.2": ("Pointer Cancellation",                     "A",   "Operable"),
    "2.5.3": ("Label in Name",                            "A",   "Operable"),
    "2.5.4": ("Motion Actuation",                         "A",   "Operable"),
    "2.5.5": ("Target Size (Enhanced)",                   "AAA", "Operable"),
    "2.5.6": ("Concurrent Input Mechanisms",              "AAA", "Operable"),
    "2.5.7": ("Dragging Movements",                       "AA",  "Operable"),
    "2.5.8": ("Target Size (Minimum)",                    "AA",  "Operable"),
    "3.1.1": ("Language of Page",                         "A",   "Understandable"),
    "3.1.2": ("Language of Parts",                        "AA",  "Understandable"),
    "3.1.3": ("Unusual Words",                            "AAA", "Understandable"),
    "3.1.4": ("Abbreviations",                            "AAA", "Understandable"),
    "3.1.5": ("Reading Level",                            "AAA", "Understandable"),
    "3.1.6": ("Pronunciation",                            "AAA", "Understandable"),
    "3.2.1": ("On Focus",                                 "A",   "Understandable"),
    "3.2.2": ("On Input",                                 "A",   "Understandable"),
    "3.2.3": ("Consistent Navigation",                    "AA",  "Understandable"),
    "3.2.4": ("Consistent Identification",                "AA",  "Understandable"),
    "3.2.5": ("Change on Request",                        "AAA", "Understandable"),
    "3.2.6": ("Consistent Help",                          "A",   "Understandable"),
    "3.3.1": ("Error Identification",                     "A",   "Understandable"),
    "3.3.2": ("Labels or Instructions",                   "A",   "Understandable"),
    "3.3.3": ("Error Suggestion",                         "AA",  "Understandable"),
    "3.3.4": ("Error Prevention (Legal, Financial, Data)","AA",  "Understandable"),
    "3.3.5": ("Help",                                     "AAA", "Understandable"),
    "3.3.6": ("Error Prevention (All)",                   "AAA", "Understandable"),
    "3.3.7": ("Redundant Entry",                          "A",   "Understandable"),
    "3.3.8": ("Accessible Authentication (Minimum)",      "AA",  "Understandable"),
    "3.3.9": ("Accessible Authentication (Enhanced)",     "AAA", "Understandable"),
    "4.1.1": ("Parsing",                                  "A",   "Robust"),
    "4.1.2": ("Name, Role, Value",                        "A",   "Robust"),
    "4.1.3": ("Status Messages",                          "AA",  "Robust"),
}

# Supplementary WCAG references for best-practice rules that have no wcagXXX axe tags.
# These are the closest related WCAG success criteria, included to give users context
# about why the issue matters. They are marked as advisory rather than strict failures.
dBpWcagRefs = {
    "accesskeys":                          ["2.1.4"],
    "aria-allowed-role":                   ["4.1.2"],
    "aria-dialog-name":                    ["4.1.2"],
    "avoid-inline-spacing":                ["1.4.12"],
    "css-orientation-lock":                ["1.3.4"],
    "empty-heading":                       ["2.4.6"],
    "empty-table-header":                  ["1.3.1"],
    "focus-trap":                          ["2.1.2"],
    "frame-tested":                        ["4.1.2"],
    "heading-order":                       ["2.4.6"],
    "hidden-content":                      ["4.1.2"],
    "identical-links-same-purpose":        ["2.4.9"],
    "label-title-only":                    ["3.3.2"],
    "landmark-banner-is-top-level":        ["1.3.1"],
    "landmark-complementary-is-top-level": ["1.3.1"],
    "landmark-contentinfo-is-top-level":   ["1.3.1"],
    "landmark-main-is-top-level":          ["1.3.1"],
    "landmark-no-duplicate-banner":        ["1.3.1"],
    "landmark-no-duplicate-contentinfo":   ["1.3.1"],
    "landmark-no-duplicate-main":          ["1.3.1"],
    "landmark-one-main":                   ["2.4.1"],
    "landmark-unique":                     ["1.3.1"],
    "meta-viewport":                       ["1.4.4"],
    "p-as-heading":                        ["1.3.1"],
    "page-has-heading-one":                ["2.4.6"],
    "presentation-role-conflict":          ["4.1.2"],
    "region":                              ["1.3.1", "2.4.1"],
    "scope-attr-valid":                    ["1.3.1"],
    "scrollable-region-focusable":         ["2.1.1"],
    "select-name":                         ["4.1.2"],
    "server-side-image-map":               ["2.1.1"],
    "skip-link":                           ["2.4.1"],
    "summary-name":                        ["4.1.2"],
    "tabindex":                            ["2.4.3"],
    "table-duplicate-name":                ["1.3.1"],
    "table-fake-caption":                  ["1.3.1"],
    "target-size":                         ["2.5.8"],
}


# --- Functions (alphabetical) ---

def flattenTarget(vTarget):
    """Convert an axe target value to a display string.
    axe target entries are strings for normal elements, or lists of strings
    for elements inside iframes or shadow DOM (path through each context)."""
    lParts = []
    vItem = None

    if isinstance(vTarget, list):
        for vItem in vTarget:
            if isinstance(vItem, list): lParts.append(" > ".join([str(v) for v in vItem]))
            else: lParts.append(str(vItem))
        return " >> ".join(lParts)
    return str(vTarget)



def buildCsvRows(dResults, dMetadata):
    dNode = {}
    dRow = {}
    dRule = {}
    iNodeCount = 0
    iRuleNodeIndex = 0
    lRows = []
    lWcagRefs = []
    sFailureSummary = ""
    sHelp = ""
    sHelpUrl = ""
    sHtml = ""
    sImpact = ""
    sOutcome = ""
    sRuleId = ""
    sStandardsRefs = ""
    sTags = ""
    sTarget = ""
    sWcagRefs = ""

    for sOutcome in aOutputSections:
        for dRule in dResults.get(sOutcome, []):
            iNodeCount = len(dRule.get("nodes", []))
            lWcagRefs = getWcagRefs(dRule)
            sHelp = str(dRule.get("help") or "")
            sHelpUrl = str(dRule.get("helpUrl") or "")
            sRuleId = str(dRule.get("id") or "")
            sStandardsRefs = " | ".join(getStandardsRefs(dRule))
            sTags = " | ".join(dRule.get("tags", []))
            sWcagRefs = " | ".join(lWcagRefs)
            if iNodeCount == 0:
                dRow = buildRowDict(dMetadata, sOutcome, dRule, sRuleId, sHelp, sHelpUrl, sTags, sWcagRefs, sStandardsRefs, 0, 0, "", "", "")
                lRows.append(dRow)
                continue
            iRuleNodeIndex = 0
            for dNode in dRule.get("nodes", []):
                iRuleNodeIndex += 1
                sFailureSummary = " ".join([str(v).strip() for v in str(dNode.get("failureSummary") or "").splitlines() if str(v).strip()])
                sHtml = str(dNode.get("html") or "")[:iCsvMaxHtmlLen]
                sImpact = str(dNode.get("impact") or dRule.get("impact") or "")
                sTarget = " | ".join([flattenTarget(v) for v in dNode.get("target", [])])
                dRow = buildRowDict(dMetadata, sOutcome, dRule, sRuleId, sHelp, sHelpUrl, sTags, sWcagRefs, sStandardsRefs, iNodeCount, iRuleNodeIndex, sTarget, sHtml, sFailureSummary, sImpact)
                lRows.append(dRow)
    lRows.sort(key=lambda dR: (dOutcomeRank.get(dR.get("outcome"), 99), dImpactRank.get(dR.get("impact"), 99), str(dR.get("ruleId") or ""), str(dR.get("target") or "")))
    return lRows


def buildConsoleSummary(dResults, dMetadata, sOutputDir):
    dRule = {}
    iCritical = 0
    iMinor = 0
    iModerate = 0
    iSerious = 0
    iTotal = 0
    lLines = []
    sPageTitle = ""
    sPageUrl = ""

    sPageTitle = str(dMetadata.get("pageTitle") or "")
    sPageUrl = str(dMetadata.get("pageUrl") or "")
    for dRule in dResults.get("violations", []):
        for dNode in dRule.get("nodes", []):
            sImpact = str(dNode.get("impact") or dRule.get("impact") or "")
            iTotal += 1
            if sImpact == "critical": iCritical += 1
            elif sImpact == "serious": iSerious += 1
            elif sImpact == "moderate": iModerate += 1
            elif sImpact == "minor": iMinor += 1
    lLines.append("")
    lLines.append(f"Page:      {sPageTitle}")
    lLines.append(f"URL:       {sPageUrl}")
    lLines.append(f"Output:    {sOutputDir}")
    lLines.append(f"Timestamp: {str(dMetadata.get('scanTimestampUtc') or '')}")
    lLines.append("")
    lLines.append(f"Violations: {iTotal} nodes across {len(dResults.get('violations', []))} rules")
    if iTotal > 0:
        lLines.append(f"  Critical: {iCritical}  Serious: {iSerious}  Moderate: {iModerate}  Minor: {iMinor}")
    return "\n".join(lLines)


def buildCheckSummaryHtml(dNode):
    """Build the How to fix HTML for a single node, following the official Deque
    pattern from doc/examples/html-handlebars.md:
      - node.any  -> Fix ANY ONE of the following (only one fix needed)
      - node.all + node.none -> Fix ALL of the following (every fix required)
    Each check may also carry relatedNodes and a data dict (e.g. contrast ratio).
    """
    dCheck = {}
    lAll = []
    lAny = []
    lNone = []
    lParts = []
    sDataExtra = ""
    sMsg = ""
    sRelated = ""

    lAny = dNode.get("any", [])
    lAll = list(dNode.get("all", [])) + list(dNode.get("none", []))

    if not lAny and not lAll:
        return ""

    def buildCheckItem(dCheck):
        sMsg = str(dCheck.get("message") or "")
        lItemParts = []
        lItemParts.append(f"<li>{html.escape(sMsg)}")
        # Render extra data inline (e.g. color-contrast ratio, expected ratio)
        vData = dCheck.get("data")
        if isinstance(vData, dict):
            lDataParts = []
            for sKey, vVal in vData.items():
                if vVal is not None and str(vVal).strip():
                    lDataParts.append(f"{html.escape(str(sKey))}: {html.escape(str(vVal))}")
            if lDataParts:
                lItemParts.append(f" <span class='check-data'>({', '.join(lDataParts)})</span>")
        elif vData is not None and str(vData).strip():
            lItemParts.append(f" <span class='check-data'>({html.escape(str(vData))})</span>")
        # Render relatedNodes if present
        lRelated = dCheck.get("relatedNodes") or []
        if lRelated:
            lItemParts.append("<ul class='related-nodes'><li><em>Related paths:</em></li>")
            for dRelNode in lRelated:
                sRelTarget = " | ".join([flattenTarget(v) for v in (dRelNode.get("target") or [])])
                sRelHtml = str(dRelNode.get("html") or "")
                if sRelTarget:
                    lItemParts.append(f"<li><code class='selector'>{html.escape(sRelTarget)}</code>")
                    if sRelHtml: lItemParts.append(f" — <code>{html.escape(sRelHtml[:200])}</code>")
                    lItemParts.append("</li>")
            lItemParts.append("</ul>")
        lItemParts.append("</li>")
        return "".join(lItemParts)

    lParts.append("<div class='fix-box'>")
    if lAny:
        lParts.append("<p class='fix-label'>Fix any one of the following:</p>")
        lParts.append("<ul>")
        for dCheck in lAny: lParts.append(buildCheckItem(dCheck))
        lParts.append("</ul>")
    if lAll:
        lParts.append("<p class='fix-label'>Fix all of the following:</p>")
        lParts.append("<ul>")
        for dCheck in lAll: lParts.append(buildCheckItem(dCheck))
        lParts.append("</ul>")
    lParts.append("</div>")
    return "".join(lParts)


def buildOutcomeSummaryRows(dResults):
    dRule = {}
    iNodeCount = 0
    lRows = []
    sOutcome = ""

    for sOutcome in aReportSections:
        iNodeCount = 0
        for dRule in dResults.get(sOutcome, []): iNodeCount += len(dRule.get("nodes", []))
        lRows.append([dSectionHeadings[sOutcome], len(dResults.get(sOutcome, [])), iNodeCount])
    return lRows



def buildNarrativeSummary(dResults, dMetadata):
    """Return an HTML string with a plain-language summary of findings at ~9th grade level."""
    dCount = {}
    iCritical = 0
    iMinor = 0
    iModerate = 0
    iSerious = 0
    iTotal = 0
    lLeadRules = []
    lNextSteps = []
    lParts = []
    lTopRules = []
    lWcagLevels = {}
    sImpact = ""
    sPageTitle = ""
    sRuleId = ""
    sWcagNote = ""

    sPageTitle = str(dMetadata.get("pageTitle") or dMetadata.get("pageUrl") or "this page")
    lViolations = dResults.get("violations", [])
    iTotal = sum(len(dRule.get("nodes", [])) for dRule in lViolations)
    iRuleCount = len(lViolations)

    if iTotal == 0:
        lParts.append("<h2 id=\"summary\">Summary and next steps</h2>")
        lParts.append("<p>No accessibility violations were detected automatically on this page. "
                      "That is a good sign, but automated tools can only find some types of problems. "
                      "A manual review by a screen reader user and a keyboard-only user will find "
                      "issues that automated scans cannot detect.</p>")
        return "\n".join(lParts)

    # Count by impact
    for dRule in lViolations:
        sImpact = str(dRule.get("impact") or "")
        n = len(dRule.get("nodes", []))
        dCount[sImpact] = dCount.get(sImpact, 0) + n
    iCritical = dCount.get("critical", 0)
    iSerious   = dCount.get("serious",  0)
    iModerate  = dCount.get("moderate", 0)
    iMinor     = dCount.get("minor",    0)

    # Top rules by instance count
    lTopRules = sorted(lViolations, key=lambda r: len(r.get("nodes", [])), reverse=True)[:3]

    # Collect unique WCAG levels touched
    for dRule in lViolations:
        for sRef in getWcagRefs(dRule):
            tInfo = getWcagScInfo(sRef)
            if tInfo:
                sLevel = tInfo[1]
                lWcagLevels[sLevel] = lWcagLevels.get(sLevel, 0) + 1

    lParts.append("<h2 id=\"summary\">Summary and next steps</h2>")

    # Opening sentence
    if iTotal == 1:
        sLead = f"One accessibility problem was found on {html.escape(sPageTitle)}."
    else:
        sLead = f"{iTotal} accessibility problems were found across {iRuleCount} rule{'s' if iRuleCount != 1 else ''} on {html.escape(sPageTitle)}."
    lParts.append(f"<p>{sLead}</p>")

    # Severity breakdown
    sSevParts = []
    if iCritical: sSevParts.append(f"<strong>{iCritical} critical</strong>")
    if iSerious:  sSevParts.append(f"<strong>{iSerious} serious</strong>")
    if iModerate: sSevParts.append(f"{iModerate} moderate")
    if iMinor:    sSevParts.append(f"{iMinor} minor")
    if sSevParts:
        lParts.append(f"<p>By severity: {', '.join(sSevParts)}. "
                      "Critical and serious problems block or significantly harm people who rely on assistive technology "
                      "such as screen readers, keyboard navigation, or voice control. Fix these first.</p>")

    # WCAG note
    if "A" in lWcagLevels:
        sWcagNote = ("Some of these problems relate to WCAG Level A criteria, "
                     "which are the minimum required for basic accessibility compliance. ")
    elif "AA" in lWcagLevels:
        sWcagNote = ("Some of these problems relate to WCAG Level AA criteria, "
                     "which most accessibility laws and policies require. ")
    if sWcagNote:
        lParts.append(f"<p>{sWcagNote}See the WCAG column in each violation card for details.</p>")

    # Top rules
    if lTopRules:
        lParts.append("<p>The most common problems found:</p><ul>")
        for dRule in lTopRules:
            sRuleId = html.escape(str(dRule.get("id") or ""))
            sHelp   = html.escape(str(dRule.get("help") or ""))
            n       = len(dRule.get("nodes", []))
            sBp     = " (best practice)" if "best-practice" in dRule.get("tags", []) else ""
            lParts.append(f"<li><a href=\"#rule-id-{sRuleId}\">{sRuleId}</a> — {sHelp}{sBp}: "
                          f"<strong>{n}</strong> instance{'s' if n != 1 else ''}</li>")
        lParts.append("</ul>")

    # Next steps
    lNextSteps = [
        "Start with critical and serious violations — these have the most impact on users with disabilities.",
        "Use the Path and Snippet details in each violation card to locate the exact element in your code.",
        "Follow the How to fix guidance for each instance, then re-run this tool to confirm the fix.",
        "After fixing automated violations, do a manual review with a screen reader (NVDA or JAWS on Windows, "
        "VoiceOver on iOS or Mac) and with keyboard-only navigation.",
        "Automated tools like this one find roughly 30 to 40 percent of accessibility issues. "
        "Manual testing and user feedback are essential for full coverage.",
    ]
    lParts.append("<p><strong>Recommended next steps:</strong></p><ol>")
    for sStep in lNextSteps: lParts.append(f"<li>{sStep}</li>")
    lParts.append("</ol>")

    return "\n".join(lParts)

def buildReportHtml(dResults, dMetadata, lRows):
    dCheck = {}
    dNode = {}
    dRule = {}
    dSummary = {}
    iCritical = 0
    iMinor = 0
    iModerate = 0
    iNodeIndex = 0
    iSerious = 0
    iTotalNodes = 0
    lCheck = []
    lFiles = []
    lImpactRows = []
    lParts = []
    lRuleLinks = []
    lRulesByFrequency = []
    lSummary = []
    lTopWcagRefs = []
    sBpTag = ""
    sCheckSection = ""
    sFixHtml = ""
    sImpactEmoji = ""
    sOutcome = ""
    sPageTitle = ""
    sRuleAnchor = ""
    sUrl = ""

    dSummary = getSummaryData(dResults, lRows)
    lImpactRows = dSummary.get("impactRows", [])
    lRulesByFrequency = dSummary.get("rulesByFrequency", [])
    lTopWcagRefs = dSummary.get("wcagRows", [])
    sPageTitle = html.escape(str(dMetadata.get("pageTitle") or sFallbackTitle))
    sUrl = html.escape(str(dMetadata.get("pageUrl") or ""))

    # Count nodes by impact for the banner
    for dRule in dResults.get("violations", []):
        for dNode in dRule.get("nodes", []):
            sImpact = str(dNode.get("impact") or dRule.get("impact") or "")
            iTotalNodes += 1
            if sImpact == "critical": iCritical += 1
            elif sImpact == "serious": iSerious += 1
            elif sImpact == "moderate": iModerate += 1
            elif sImpact == "minor": iMinor += 1

    lFiles = [
        f"<li><a href=\"{html.escape(sReportWorkbookName)}\">{html.escape(sReportWorkbookName)}</a> — Excel workbook</li>",
        f"<li><a href=\"{html.escape(sCsvName)}\">{html.escape(sCsvName)}</a> — spreadsheet of violations</li>",
        f"<li><a href=\"{html.escape(sJsonName)}\">{html.escape(sJsonName)}</a> — full raw data</li>",
        f"<li><a href=\"{html.escape(sSourceName)}\">{html.escape(sSourceName)}</a> — saved page source</li>",
        f"<li><a href=\"{html.escape(sScreenshotName)}\">{html.escape(sScreenshotName)}</a> — page screenshot</li>",
    ]

    # CSS
    sStyle = (
        "body{font-family:Segoe UI,Arial,sans-serif;line-height:1.6;margin:0;padding:0;background:#f5f5f5;color:#1a1a1a}"
        "main{max-width:960px;margin:0 auto;padding:1.5rem}"
        "a{color:#0060b9;text-decoration:underline}"
        "a:hover{text-decoration:none}"
        "h1{font-size:1.6rem;margin:0 0 .25rem 0}"
        "h2{font-size:1.25rem;border-bottom:2px solid #0060b9;padding-bottom:.25rem;margin-top:2rem}"
        "h3{font-size:1.1rem;margin-top:1.5rem}"
        "h4{font-size:1rem;margin:.75rem 0 .25rem 0}"
        "h5{font-size:.95rem;margin:.5rem 0 .2rem 0;color:#444}"
        "h6{font-size:.9rem;margin:.4rem 0 .15rem 0;color:#555}"
        "code,pre{font-family:Consolas,'Courier New',monospace;font-size:.88em}"
        "pre{background:#1e1e1e;color:#d4d4d4;padding:.75rem 1rem;border-radius:4px;overflow-x:auto;white-space:pre-wrap;word-break:break-all}"
        "table{border-collapse:collapse;width:100%;margin:.5rem 0 1rem 0;background:#fff;border-radius:4px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,.1)}"
        "th,td{border:1px solid #ddd;padding:.5rem .75rem;text-align:left;vertical-align:top}"
        "th{background:#0060b9;color:#fff;font-weight:600}"
        "tr:nth-child(even){background:#f9f9f9}"
        "nav ul{padding-left:1.25rem;margin:.25rem 0}"
        "nav li{margin:.15rem 0}"
        ".banner{background:#0060b9;color:#fff;padding:1rem 1.5rem;margin-bottom:1rem}"
        ".banner h1{color:#fff}"
        ".banner p{margin:.2rem 0;font-size:.95rem;opacity:.9}"
        ".impact-bar{display:flex;gap:.5rem;flex-wrap:wrap;margin:.5rem 0}"
        ".badge{display:inline-block;padding:.2rem .6rem;border-radius:3px;font-weight:700;font-size:.85rem;color:#fff}"
        ".badge-critical{background:#c00}"
        ".badge-serious{background:#c55000}"
        ".badge-moderate{background:#856404}"
        ".badge-minor{background:#555}"
        ".badge-pass{background:#1a7a1a}"
        ".rule-card{background:#fff;border-left:5px solid #999;border-radius:4px;margin:1rem 0;padding:1rem 1.25rem;box-shadow:0 1px 3px rgba(0,0,0,.08)}"
        ".rule-card.critical{border-color:#c00}"
        ".rule-card.serious{border-color:#c55000}"
        ".rule-card.moderate{border-color:#856404}"
        ".rule-card.minor{border-color:#555}"
        ".rule-meta{font-size:.88rem;color:#555;margin:.2rem 0}"
        ".node-card{background:#f8f8f8;border:1px solid #ddd;border-radius:3px;margin:.75rem 0;padding:.75rem 1rem}"
        ".node-label{font-size:.8rem;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:#666;margin-bottom:.2rem}"
        ".selector{font-family:Consolas,'Courier New',monospace;font-size:.85em;background:#eee;padding:.15rem .4rem;border-radius:2px;word-break:break-all}"
        ".fix-box{background:#fff8e1;border:1px solid #f0c040;border-radius:3px;padding:.5rem .75rem;margin:.5rem 0}"
        ".fix-box p{margin:.2rem 0;font-size:.9rem}"
        ".fix-label{font-weight:700;color:#6b4c00;font-size:.85rem;text-transform:uppercase;letter-spacing:.05em}"
        ".back-link{font-size:.85rem;float:right;margin-top:.2rem}"
        ".check-data{color:#555;font-size:.85em}"
        ".related-nodes{margin:.3rem 0 0 .5rem;font-size:.85em;list-style:disc;padding-left:1.2rem}"
        ".muted{color:#666;font-size:.9rem}"
        "dl.meta-grid{display:grid;grid-template-columns:max-content 1fr;gap:.2rem .75rem;font-size:.9rem}"
        "dt{font-weight:600}"
        "dd{margin:0 0 .2rem 0;word-break:break-all}"
        "#skip-link{position:absolute;left:-999px}"
        "#skip-link:focus{left:1rem;top:1rem;z-index:9999;background:#fff;padding:.5rem 1rem;border:2px solid #0060b9}"
        "details{margin:.5rem 0}"
        "summary{cursor:pointer;font-weight:600;padding:.4rem .5rem;background:#f0f0f0;border-radius:3px;user-select:none}"
        "summary:hover{background:#e0e0e0}"
        ".badge-bp{background:#5a3e8a}"
        ".tag-bp{display:inline-block;padding:.1rem .45rem;border-radius:3px;font-size:.78rem;font-weight:700;background:#ede7f6;color:#4a235a;margin-left:.4rem;vertical-align:middle}"
    )

    lParts.append("<!doctype html>")
    lParts.append("<html lang=\"en\">")
    lParts.append("<head>")
    lParts.append("<meta charset=\"utf-8\">")
    lParts.append("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">")
    lParts.append(f"<title>Accessibility report — {sPageTitle}</title>")
    lParts.append(f"<style>{sStyle}</style>")
    lParts.append("</head>")
    lParts.append("<body>")
    lParts.append("<a href=\"#main-content\" id=\"skip-link\">Skip to main content</a>")

    # Banner
    lParts.append("<header class=\"banner\">")
    lParts.append(f"<h1>Accessibility report — {sPageTitle}</h1>")
    lParts.append(f"<p><a href=\"{sUrl}\" style=\"color:#cde\">{sUrl}</a></p>")
    lParts.append(f"<p>Scanned: {html.escape(str(dMetadata.get('scanTimestampUtc') or ''))} &nbsp;|&nbsp; {html.escape(str(dMetadata.get('browserVersion') or ''))}</p>")
    lParts.append("</header>")

    lParts.append("<main id=\"main-content\">")

    # Quick summary box
    iViolRules = len(dResults.get("violations", []))
    lParts.append("<section aria-label=\"Quick summary\">")
    lParts.append(f"<p><strong>{iTotalNodes} failed instance{'' if iTotalNodes==1 else 's'}</strong> across <strong>{iViolRules} rule{'' if iViolRules==1 else 's'}</strong> with accessibility violations.</p>")
    if iTotalNodes > 0:
        lParts.append("<div class=\"impact-bar\" aria-label=\"Counts by impact\">")
        if iCritical: lParts.append(f"<span class=\"badge badge-critical\">Critical: {iCritical}</span>")
        if iSerious:  lParts.append(f"<span class=\"badge badge-serious\">Serious: {iSerious}</span>")
        if iModerate: lParts.append(f"<span class=\"badge badge-moderate\">Moderate: {iModerate}</span>")
        if iMinor:    lParts.append(f"<span class=\"badge badge-minor\">Minor: {iMinor}</span>")
        lParts.append("</div>")
        lParts.append("<p class=\"muted\"><strong>Critical</strong> — blocks access completely. <strong>Serious</strong> — very hard to use. <strong>Moderate</strong> — causes difficulty. <strong>Minor</strong> — small problem with a workaround.</p>")
    lParts.append("</section>")

    # Table of contents
    lParts.append("<h2 id=\"table-of-contents\">Table of contents</h2>")
    lParts.append("<nav aria-label=\"Table of contents\"><ul>")
    lParts.append("<li><a href=\"#scan-details\">Scan details</a></li>")
    lParts.append("<li><a href=\"#patterns\">Patterns</a></li>")
    lParts.append("<li><a href=\"#violations\">Violations</a><ul>")
    for iRuleIdx, dRule in enumerate(dResults.get("violations", []), 1):
        sRuleAnchor = f"rule-{iRuleIdx}"
        sRuleLabel = html.escape(str(dRule.get("id") or ""))
        sRuleHelp = html.escape(str(dRule.get("help") or ""))
        sImpact = html.escape(str(dRule.get("impact") or ""))
        sImpactEmoji = dImpactEmoji.get(str(dRule.get("impact") or ""), "")
        sBpTagToc = " <span class=\"tag-bp\">Best Practice</span>" if "best-practice" in dRule.get("tags", []) else ""
        lParts.append(f"<li><a href=\"#{html.escape(sRuleAnchor)}\">{sImpactEmoji} {sRuleLabel}: {sRuleHelp}</a> <span class=\"badge badge-{sImpact}\" style=\"font-size:.75rem\">{sImpact}</span>{sBpTagToc}</li>")
    lParts.append("</ul></li>")
    lParts.append("<li><a href=\"#summary\">Summary and next steps</a></li>")
    lParts.append("<li><a href=\"#output-files\">Output files</a></li>")
    lParts.append("<li><a href=\"#glossary\">Glossary</a></li>")
    lParts.append("<li><a href=\"#resources\">Resources</a></li>")
    lParts.append("</ul></nav>")

    # Scan details
    lParts.append("<h2 id=\"scan-details\">Scan details</h2>")
    lParts.append("<dl class=\"meta-grid\">")
    for sLabel, sValue in [
        ["URL scanned", str(dMetadata.get("pageUrl") or "")],
        ["Page title", str(dMetadata.get("pageTitle") or "")],
        ["Scan time (UTC)", str(dMetadata.get("scanTimestampUtc") or "")],
        ["Browser", str(dMetadata.get("browserVersion") or "")],
        ["Viewport", f"{dMetadata.get('viewportWidth') or 0} \u00d7 {dMetadata.get('viewportHeight') or 0} pixels"],
        ["Testing engine", str(dMetadata.get("axeSource") or "")],
        ["User agent", str(dMetadata.get("userAgent") or "")],
    ]:
        lParts.append(f"<dt>{html.escape(sLabel)}</dt><dd>{html.escape(sValue)}</dd>")
    lParts.append("</dl>")
    lParts.append("<h3>All outcome counts</h3>")
    lParts.append("<table><thead><tr><th>Outcome</th><th>Rules</th><th>Instances</th><th>What it means</th></tr></thead><tbody>")
    dOutcomeExplain = {
        "Violations": "Confirmed accessibility problems that need to be fixed.",
        "Needs Review": "Possible problems that need a person to check manually.",
        "Passes": "Rules that this page passed automatically.",
        "Inapplicable": "Rules that do not apply to this page.",
    }
    for lRow in buildOutcomeSummaryRows(dResults):
        sLabel = str(lRow[0])
        lParts.append(f"<tr><td><strong>{html.escape(sLabel)}</strong></td><td>{html.escape(str(lRow[1]))}</td><td>{html.escape(str(lRow[2]))}</td><td class=\"muted\">{html.escape(dOutcomeExplain.get(sLabel,''))}</td></tr>")
    lParts.append("</tbody></table>")

    # Patterns
    lParts.append("<h2 id=\"patterns\">Patterns</h2>")
    lParts.append("<p class=\"muted\">These tables show which problems appear most often, to help you decide where to start.</p>")
    lParts.append("<h3>Failed instances by impact</h3>")
    lParts.append("<table><thead><tr><th>Impact</th><th>Instances</th><th>What it means</th></tr></thead><tbody>")
    dImpactExplain = {
        "critical": "Blocks some users completely. Fix immediately.",
        "serious": "Very hard for some users to work around. High priority.",
        "moderate": "Causes real difficulty. Fix when possible.",
        "minor": "Small issue. A workaround usually exists.",
    }
    for lRow in lImpactRows:
        sImp = str(lRow[0])
        lParts.append(f"<tr><td><span class=\"badge badge-{html.escape(sImp)}\">{html.escape(sImp)}</span></td><td>{html.escape(str(lRow[1]))}</td><td class=\"muted\">{html.escape(dImpactExplain.get(sImp,''))}</td></tr>")
    if not lImpactRows: lParts.append("<tr><td colspan=\"3\">No violations found.</td></tr>")
    lParts.append("</tbody></table>")
    lParts.append("<h3>Most common rules by instance count</h3>")
    if lRulesByFrequency:
        lParts.append("<table><thead><tr><th>Rule</th><th>Failing elements</th></tr></thead><tbody>")
        for lRow in lRulesByFrequency:
            lParts.append(f"<tr><td><a href=\"#rule-id-{html.escape(str(lRow[0]))}\">{html.escape(str(lRow[0]))}</a></td><td>{html.escape(str(lRow[1]))}</td></tr>")
        lParts.append("</tbody></table>")
    else:
        lParts.append("<p>No violations found.</p>")
    lParts.append("<h3>Most common WCAG criteria</h3>")
    if lTopWcagRefs:
        lParts.append("<table><thead><tr><th>SC</th><th>Name</th><th>Level</th><th>Principle</th><th>Instances</th></tr></thead><tbody>")
        for lRow in lTopWcagRefs:
            sRef = str(lRow[0])
            tInfo = getWcagScInfo(sRef)
            sName = html.escape(tInfo[0]) if tInfo else ""
            sLevel = html.escape(tInfo[1]) if tInfo else ""
            sPrinciple = html.escape(tInfo[2]) if tInfo else ""
            lParts.append(f"<tr><td><a href=\"{html.escape(sWcagBaseUrl)}{html.escape(sRef)}/\">{html.escape(sRef)}</a></td><td>{sName}</td><td>{sLevel}</td><td>{sPrinciple}</td><td>{html.escape(str(lRow[1]))}</td></tr>")
        lParts.append("</tbody></table>")
    else:
        lParts.append("<p class=\"muted\">No violations with WCAG SC references were detected. Best-practice violations may still relate to WCAG criteria — see the per-rule WCAG notes in the Violations section below.</p>")

    # Violations
    lParts.append("<h2 id=\"violations\">Violations</h2>")
    lParts.append("<p class=\"muted\">Each section below is one failing rule. Inside each rule, every element on the page that failed is listed with its HTML, CSS path, and specific fix guidance.</p>")
    if not dResults.get("violations"):
        lParts.append("<p>No violations were detected on this page.</p>")
    else:
        for iRuleIdx, dRule in enumerate(dResults.get("violations", []), 1):
            sRuleAnchor = f"rule-{iRuleIdx}"
            sRuleId = str(dRule.get("id") or "")
            sImpact = str(dRule.get("impact") or "")
            iNodeCount = len(dRule.get("nodes", []))
            lRuleLinks = getRuleLinks(dRule)
            lWcagRefs = getWcagRefs(dRule)
            lParts.append(f"<section class=\"rule-card {html.escape(sImpact)}\" id=\"{html.escape(sRuleAnchor)}\" aria-label=\"Rule: {html.escape(sRuleId)}\">")
            lParts.append(f"<a href=\"#table-of-contents\" class=\"back-link\" aria-label=\"Back to table of contents\">&uarr; Contents</a>")
            sImpactEmoji = dImpactEmoji.get(sImpact, "")
            lParts.append(f"<h3 id=\"rule-id-{html.escape(sRuleId)}\">{sImpactEmoji} {html.escape(sRuleId)}: {html.escape(str(dRule.get('help') or ''))}</h3>")
            sBpTag = "<span class=\"tag-bp\">Best Practice</span>" if "best-practice" in dRule.get("tags", []) else ""
            lParts.append(f"<p class=\"rule-meta\"><span class=\"badge badge-{html.escape(sImpact)}\">{html.escape(sImpact)}</span>{sBpTag} &nbsp; {iNodeCount} failed instance{'' if iNodeCount==1 else 's'}</p>")
            lParts.append(f"<p>{html.escape(str(dRule.get('description') or ''))}</p>")
            if lWcagRefs:
                bIsBp = "best-practice" in dRule.get("tags", []) and not any(getWcagRef(t) for t in dRule.get("tags", []))
                lWcagParts = []
                for sRef in lWcagRefs:
                    tInfo = getWcagScInfo(sRef)
                    sName = f" — {html.escape(tInfo[0])} (Level {html.escape(tInfo[1])})" if tInfo else ""
                    lWcagParts.append(f"<a href=\"{html.escape(sWcagBaseUrl)}{html.escape(sRef)}/\">{html.escape(sRef)}</a>{sName}")
                sWcagLinks = "<br>".join(lWcagParts)
                sAdvisory = " <span class=\"muted\" style=\"font-style:italic\">(advisory — related criteria for this best-practice rule)</span>" if bIsBp else ""
                lParts.append(f"<p class=\"rule-meta\"><strong>WCAG:</strong>{sAdvisory}<br>{sWcagLinks} &nbsp;|&nbsp; <a href=\"{html.escape(str(dRule.get('helpUrl') or ''))}\">Deque rule documentation</a></p>")
            else:
                lParts.append(f"<p class=\"rule-meta\"><a href=\"{html.escape(str(dRule.get('helpUrl') or ''))}\">Deque rule documentation</a></p>")

            # Failing elements
            iNodeIndex = 0
            lParts.append(f"<details open><summary>{iNodeCount} failed instance{"" if iNodeCount==1 else "s"} — select to expand or collapse</summary>")
            for dNode in dRule.get("nodes", []):
                iNodeIndex += 1
                sNodeImpact = str(dNode.get("impact") or sImpact)
                sTarget = " | ".join([flattenTarget(v) for v in dNode.get("target", [])])
                sHtmlSnippet = str(dNode.get("html") or "")
                sFailSummary = str(dNode.get("failureSummary") or "")

                lParts.append("<div class=\"node-card\">")
                lParts.append(f"<p class=\"node-label\">Instance {iNodeIndex} of {iNodeCount}</p>")

                # CSS path
                if sTarget:
                    lParts.append("<p class=\"muted\" style=\"margin:.2rem 0 0\"><strong>Path</strong> — CSS selector for this element:</p>")
                    lParts.append(f"<p><code class=\"selector\">{html.escape(sTarget)}</code></p>")

                # HTML snippet
                if sHtmlSnippet:
                    lParts.append("<p class=\"muted\" style=\"margin:.4rem 0 0\"><strong>Snippet</strong> — HTML of this element:</p>")
                    lParts.append(f"<pre><code>{html.escape(sHtmlSnippet)}</code></pre>")

                # How to fix — use proper any/all/none semantics per Deque handlebars example
                sFixHtml = buildCheckSummaryHtml(dNode)
                if sFixHtml:
                    lParts.append(sFixHtml)
                elif sFailSummary:
                    lParts.append("<div class=\"fix-box\">")
                    lParts.append("<p class=\"fix-label\">How to fix</p>")
                    lParts.append(f"<pre>{html.escape(sFailSummary)}</pre>")
                    lParts.append("</div>")

                # Impact at node level if different from rule
                if sNodeImpact and sNodeImpact != sImpact:
                    lParts.append(f"<p class=\"muted\">Element impact: <span class=\"badge badge-{html.escape(sNodeImpact)}\">{html.escape(sNodeImpact)}</span></p>")

                lParts.append("</div>")
            lParts.append("</details>")
            lParts.append("</section>")

    # Narrative summary
    lParts.append(buildNarrativeSummary(dResults, dMetadata))

    # Output files
    lParts.append("<h2 id=\"output-files\">Output files</h2>")
    lParts.append("<p class=\"muted\">All files are saved in the same folder as this report.</p>")
    lParts.append(f"<ul>{''.join(lFiles)}</ul>")

    # Glossary
    lParts.append("<h2 id=\"glossary\">Glossary</h2>")
    lParts.append("<dl>")
    for lRow in aGlossaryRows: lParts.append(f"<dt><strong>{html.escape(str(lRow[0]))}</strong></dt><dd>{html.escape(str(lRow[1]))}</dd>")
    lParts.append("</dl>")

    # Resources
    lParts.append("<h2 id=\"resources\">Resources</h2>")
    lParts.append("<ul>")
    lParts.append(f"<li><a href=\"{html.escape(sWcagBaseUrl)}\">WCAG 2.2 Understanding — what each criterion means and how to meet it</a></li>")
    lParts.append(f"<li><a href=\"{html.escape(sAccessibilityInsightsUrl)}\">Accessibility Insights — free browser extension for manual testing</a></li>")
    lParts.append(f"<li><a href=\"{html.escape(sMsAccessibilityUrl)}\">Microsoft Accessibility — guidance and tools from Microsoft</a></li>")
    lParts.append("<li><a href=\"https://dequeuniversity.com/rules/axe/4.11/\">Deque — full list of axe-core rules with explanations</a></li>")
    lParts.append("</ul>")

    lParts.append("</main>")
    lParts.append("</body>")
    lParts.append("</html>")
    return "\n".join(lParts)


def buildRowDict(dMetadata, sOutcome, dRule, sRuleId, sHelp, sHelpUrl, sTags, sWcagRefs, sStandardsRefs, iNodeCount, iNodeIndex, sTarget, sHtml, sFailureSummary, sImpact=""):
    return {
        "axeSource": str(dMetadata.get("axeSource") or ""),
        "browserVersion": str(dMetadata.get("browserVersion") or ""),
        "description": str(dRule.get("description") or ""),
        "failureSummary": sFailureSummary,
        "help": sHelp,
        "helpUrl": sHelpUrl,
        "html": sHtml,
        "impact": sImpact or str(dRule.get("impact") or ""),
        "outcome": sOutcome,
        "pageTitle": str(dMetadata.get("pageTitle") or ""),
        "pageUrl": str(dMetadata.get("pageUrl") or ""),
        "ruleId": sRuleId,
        "ruleNodeCount": iNodeCount,
        "ruleNodeIndex": iNodeIndex,
        "scanTimestampUtc": str(dMetadata.get("scanTimestampUtc") or ""),
        "standardsRefs": sStandardsRefs,
        "tags": sTags,
        "target": sTarget,
        "wcagRefs": sWcagRefs,
    }


def chooseOutputDir(pathBaseDir, sPageTitle):
    iIndex = 0
    pathCandidate = None
    sBaseName = ""
    sName = ""

    sBaseName = getSafeTitle(sPageTitle)
    pathCandidate = pathBaseDir / sBaseName
    if not pathCandidate.exists():
        pathCandidate.mkdir(parents=True, exist_ok=False)
        return pathCandidate
    iIndex = 1
    while True:
        sName = f"{sBaseName}-{iIndex:0{iSuffixDigits}d}"
        pathCandidate = pathBaseDir / sName
        if not pathCandidate.exists():
            pathCandidate.mkdir(parents=True, exist_ok=False)
            return pathCandidate
        iIndex += 1


def ensureSuccess(bCondition, sMessage):
    if not bCondition: raise RuntimeError(sMessage)
    return True


def fetchText(sUrl):
    dHeaders = {}
    oRequest = None
    oResponse = None
    sText = ""

    dHeaders = {"User-Agent": sUserAgent}
    oRequest = urllib.request.Request(sUrl, headers=dHeaders)
    with urllib.request.urlopen(oRequest, timeout=iCdnTimeoutSec) as oResponse: sText = oResponse.read().decode("utf-8")
    return sText


def getAxeScript(page):
    oError = None
    sAxeScript = ""
    sUrl = ""

    for sUrl in aAxeCdnUrls:
        try:
            page.add_script_tag(url=sUrl)
            return sUrl
        except Exception as oError:
            continue
    for sUrl in aAxeCdnUrls:
        try:
            sAxeScript = fetchText(sUrl)
            page.add_script_tag(content=sAxeScript)
            return sUrl
        except Exception as oError:
            continue
    raise RuntimeError("Unable to load axe-core from the configured CDN URLs.")


def getImpactRows(lRows):
    dCounts = {}
    dRow = {}
    lOutput = []
    sImpact = ""

    for dRow in lRows:
        sImpact = str(dRow.get("impact") or "")
        if not sImpact: continue
        dCounts[sImpact] = int(dCounts.get(sImpact, 0)) + 1
    for sImpact in sorted(dCounts.keys(), key=lambda sK: dImpactRank.get(sK, 99)): lOutput.append([sImpact, dCounts[sImpact]])
    return lOutput


def getNormalizedUrl(sInput):
    """Convert a user-supplied input to a fully qualified URL or local file URI.
    Accepts bare domains (microsoft.com), IP addresses, paths with or without
    scheme, and local HTML files. Unrecognised inputs are returned unchanged."""
    oPath = None
    sCandidate = ""
    sSuffix = ""
    sLower = ""

    sCandidate = sInput.strip().strip('"')
    sLower = sCandidate.lower()
    # Already has a scheme — pass through as-is
    if "://" in sLower: return sCandidate
    # Existing local HTML file
    oPath = pathlib.Path(sCandidate).expanduser().resolve()
    if oPath.exists() and oPath.is_file() and oPath.suffix.lower() in aAllowedLocalExtensions:
        return oPath.as_uri()
    # Existing non-HTML file — treat as URL list; do not URL-ify
    if oPath.exists() and oPath.is_file(): return sCandidate
    # Has a recognised local file extension and no path separator — treat as local filename
    sSuffix = pathlib.Path(sCandidate.split("?")[0].split("#")[0]).suffix.lower()
    if sSuffix in aLocalFileExts and "/" not in sCandidate: return sCandidate
    # Looks like an IP address (optional port/path)
    if reIpAddress.match(sCandidate): return f"https://{sCandidate}"
    # Looks like a domain name (optional port/path)
    if reDomain.match(sCandidate): return f"https://{sCandidate}"
    return sCandidate


def getPageSnapshot(page, sPageUrl):
    dCss = {}
    dScript = {}
    sBaseUrl = ""
    sHref = ""
    sHtml = ""
    sInlineHtml = ""
    sSrc = ""

    sHtml = str(page.content())
    try:
        sBaseUrl = str(page.url or sPageUrl)
        for sHref in page.eval_on_selector_all("link[rel='stylesheet'][href]", "elements => elements.map(e => e.getAttribute('href'))"):
            try: dCss[urllib.parse.urljoin(sBaseUrl, str(sHref))] = fetchText(urllib.parse.urljoin(sBaseUrl, str(sHref)))
            except Exception: continue
        for sSrc in page.eval_on_selector_all("script[src]", "elements => elements.map(e => e.getAttribute('src'))"):
            try: dScript[urllib.parse.urljoin(sBaseUrl, str(sSrc))] = fetchText(urllib.parse.urljoin(sBaseUrl, str(sSrc)))
            except Exception: continue
    except Exception:
        return sHtml
    sInlineHtml = sHtml
    for sHref in list(dCss.keys()):
        sCssContent = f'<style data-pagecheck-source="{html.escape(sHref)}">\n{dCss[sHref]}\n</style>'
        sInlineHtml = re.sub(rf'<link\b[^>]*href=["\']{re.escape(sHref)}["\'][^>]*>', lambda m, s=sCssContent: s, sInlineHtml, flags=re.IGNORECASE)
    for sSrc in list(dScript.keys()):
        sScriptContent = f'<script data-pagecheck-source="{html.escape(sSrc)}">\n{dScript[sSrc]}\n</script>'
        sInlineHtml = re.sub(rf'<script\b[^>]*src=["\']{re.escape(sSrc)}["\'][^>]*>\s*</script>', lambda m, s=sScriptContent: s, sInlineHtml, flags=re.IGNORECASE)
    return sInlineHtml


def getRuleFrequencyRows(lRows):
    dCounts = {}
    dRow = {}
    lOutput = []
    sRuleId = ""

    for dRow in lRows:
        sRuleId = str(dRow.get("ruleId") or "")
        if not sRuleId: continue
        dCounts[sRuleId] = int(dCounts.get(sRuleId, 0)) + 1
    for sRuleId in sorted(dCounts.keys(), key=lambda sK: (-int(dCounts.get(sK, 0)), sK)): lOutput.append([sRuleId, dCounts[sRuleId]])
    return lOutput[:20]


def getRuleLinks(dRule):
    lLinks = []
    sRuleId = ""
    sTag = ""
    sWcagRef = ""

    sRuleId = html.escape(str(dRule.get("id") or ""))
    lLinks.append(f"<li><a href=\"{sMsAccessibilityUrl}\">Microsoft Accessibility guidance</a> for rule <code>{sRuleId}</code></li>")
    for sTag in dRule.get("tags", []):
        if not any(str(sTag).startswith(sPrefix) for sPrefix in aRelevantTagPrefixes): continue
        sWcagRef = getWcagRef(str(sTag))
        if sWcagRef:
            lLinks.append(f"<li><a href=\"{sWcagBaseUrl}{html.escape(sWcagRef)}/\">WCAG Understanding: {html.escape(sWcagRef)}</a></li>")
            continue
        lLinks.append(f"<li>{html.escape(str(sTag))}</li>")
    if len(lLinks) == 1: lLinks.append(f"<li><a href=\"{sAccessibilityInsightsUrl}\">Accessibility Insights overview</a></li>")
    return lLinks


def getSafeTitle(sTitle):
    sName = ""

    sName = re.sub(r"\s+", "-", str(sTitle or sFallbackTitle).strip().lower())
    sName = re.sub(r"[^a-z0-9._-]", "-", sName)
    sName = re.sub(r"-+", "-", sName).strip("-._")
    if not sName: sName = sFallbackTitle
    sName = sName[:iMaxTitleLen].strip("-._")
    if not sName: sName = sFallbackTitle
    return sName


def getStandardsRefs(dRule):
    lRefs = []
    sTag = ""

    for sTag in dRule.get("tags", []):
        if not any(str(sTag).startswith(sPrefix) for sPrefix in aRelevantTagPrefixes): continue
        if str(sTag).startswith("wcag"): continue
        lRefs.append(str(sTag))
    return lRefs


def getSummaryData(dResults, lRows):
    dSummary = {}

    dSummary = {
        "impactRows": getImpactRows(lRows),
        "outcomeRows": buildOutcomeSummaryRows(dResults),
        "rulesByFrequency": getRuleFrequencyRows(lRows),
        "wcagRows": getWcagFrequencyRows(lRows),
    }
    return dSummary


def getUrlsFromFile(sInput):
    """Read a text file and return a list of non-blank URLs, one per line.
    Raises FileNotFoundError or PermissionError if the file cannot be read.
    Each line is stripped; lines that are blank or start with # are ignored."""
    lUrls = []
    oPath = None
    sLine = ""

    oPath = pathlib.Path(sInput).expanduser().resolve()
    with oPath.open("r", encoding="utf-8", errors="replace") as oFile:
        for sLine in oFile:
            sLine = sLine.strip()
            if not sLine or sLine.startswith("#"): continue
            lUrls.append(sLine)
    if not lUrls: raise ValueError(f"No URLs found in file: {sInput}")
    return lUrls


def getWcagFrequencyRows(lRows):
    dCounts = {}
    dRow = {}
    lOutput = []
    sPart = ""
    sWcagRefs = ""

    for dRow in lRows:
        sWcagRefs = str(dRow.get("wcagRefs") or "")
        if not sWcagRefs: continue
        for sPart in [sPart.strip() for sPart in sWcagRefs.split("|") if sPart.strip()]:
            dCounts[sPart] = int(dCounts.get(sPart, 0)) + 1
    for sPart in sorted(dCounts.keys(), key=lambda sK: (-int(dCounts.get(sK, 0)), sK)): lOutput.append([sPart, dCounts[sPart]])
    return lOutput[:20]


def getWcagRef(sTag):
    """Return a numeric WCAG SC ref like '1.4.3' from an axe tag like 'wcag143',
    or empty string if the tag is not a specific SC reference."""
    sDigits = ""
    sRef = ""

    if not str(sTag).startswith("wcag"): return ""
    if re.fullmatch(r"wcag\d+[a-z]+", str(sTag)): return ""
    sDigits = re.sub(r"[^0-9]", "", str(sTag))
    if len(sDigits) < 3: return ""
    sRef = f"{sDigits[0]}.{sDigits[1]}.{sDigits[2]}"
    return sRef


def getWcagRefs(dRule):
    lRefs = []
    sRef = ""
    sRuleId = ""
    sTag = ""

    # Primary: derive refs from axe wcagXXX tags
    for sTag in dRule.get("tags", []):
        sRef = getWcagRef(str(sTag))
        if sRef and sRef not in lRefs: lRefs.append(sRef)
    # Supplementary: if rule is best-practice with no wcag tags, use advisory map
    if not lRefs and "best-practice" in dRule.get("tags", []):
        sRuleId = str(dRule.get("id") or "")
        for sRef in dBpWcagRefs.get(sRuleId, []):
            if sRef not in lRefs: lRefs.append(sRef)
    return lRefs


def getWcagScInfo(sRef):
    """Return (shortName, level, principle) for a WCAG SC number, or None if unknown."""
    return dWcagSc.get(sRef)


def isUrlListFile(sInput):
    """Return True if sInput resolves to an existing file that is not an
    allowed HTML extension. Such a file is treated as a URL list."""
    oPath = None

    oPath = pathlib.Path(sInput).expanduser().resolve()
    if not oPath.exists() or not oPath.is_file(): return False
    if oPath.suffix.lower() in aAllowedLocalExtensions: return False
    return True


def parseArguments():
    argParser = None

    argParser = argparse.ArgumentParser(
        prog=sProgramName,
        description=(
            "Open a web page or local HTML file in Microsoft Edge, run axe-core, "
            "and write structured accessibility outputs. "
            "Pass a text file containing one URL per line to scan multiple pages in sequence."
        ),
        epilog=(
            f"Single URL:  {sProgramName} https://example.com\n"
            f"Local file:  {sProgramName} C:\\work\\sample.html\n"
            f"URL list:    {sProgramName} urls.txt"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    argParser.add_argument(
        "inputValue",
        nargs="?",
        help=(
            "URL, local HTML file, or path to a text file containing one URL per line. "
            "If a text file is given, each URL is scanned in sequence and report.html "
            "is not opened automatically."
        ),
    )
    argParser.add_argument(
        "--wait",
        type=int,
        default=0,
        metavar="SECONDS",
        help="Extra seconds to wait after the page loads before running axe. Useful for JS-heavy pages. Default: 0.",
    )
    argParser.add_argument("-v", "--version", action="version", version=f"%(prog)s {sProgramVersion}")
    return argParser.parse_args()


def scanUrl(sInput, sNormalizedUrl, browser, context, pathBaseDir, bOpenReport, iExtraWaitMs=0):
    """Run a single-URL scan. Returns the output directory path string, or raises."""
    dMetadata = {}
    dResults = {}
    lArgs = []
    lRows = []
    oPage = None
    pathOutputDir = None
    sAxeSource = ""
    sPageTitle = ""
    sResultsJson = ""
    sSnapshot = ""

    try:
        oPage = context.new_page()
        print(f"  Navigating to {sNormalizedUrl} ...")
        oPage.goto(sNormalizedUrl, timeout=iDefaultNavTimeoutMs, wait_until="load")
        # After the load event, wait briefly for network to settle.
        # A short networkidle timeout lets SPA content finish rendering on most sites
        # while sites with persistent connections (e.g. WebSocket-heavy SPAs) simply
        # time out and continue after iNetworkIdleTimeoutMs milliseconds.
        try:
            oPage.wait_for_load_state("networkidle", timeout=iNetworkIdleTimeoutMs)
        except Exception:
            pass
        oPage.wait_for_timeout(iDefaultPostLoadDelayMs)
        # Scroll to bottom and back to trigger lazy-loaded content, then wait briefly
        oPage.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        oPage.wait_for_timeout(500)
        oPage.evaluate("window.scrollTo(0, 0)")
        oPage.wait_for_timeout(500)
        if iExtraWaitMs > 0:
            print(f"  Waiting {iExtraWaitMs // 1000}s extra ...")
            oPage.wait_for_timeout(iExtraWaitMs)
        sPageTitle = str(oPage.title() or sFallbackTitle)
        pathOutputDir = chooseOutputDir(pathBaseDir, sPageTitle)
        print(f"  Output:    {pathOutputDir}")
        print("  Running axe-core ...")
        sAxeSource = getAxeScript(oPage)
        ensureSuccess(bool(oPage.evaluate("() => Boolean(window.axe && window.axe.run)")), "axe-core did not load into the page.")
        sResultsJson = oPage.evaluate("async (opts) => JSON.stringify(await window.axe.run(document, opts))", aAxeRunOptions)
        dResults = json.loads(sResultsJson)
        dMetadata = {
            "axeSource": sAxeSource,
            "browserChannel": sBrowserChannel,
            "browserVersion": str(browser.version),
            "inputValue": sInput,
            "navTimeoutMs": iDefaultNavTimeoutMs,
            "normalizedUrl": sNormalizedUrl,
            "pageTitle": sPageTitle,
            "pageUrl": str(oPage.url or sNormalizedUrl),
            "extraWaitMs": iExtraWaitMs,
            "postLoadDelayMs": iDefaultPostLoadDelayMs,
            "programName": sProgramName,
            "programVersion": sProgramVersion,
            "scanTimestampUtc": datetime.datetime.now(datetime.timezone.utc).replace(microsecond=0).isoformat(),
            "userAgent": str(oPage.evaluate("() => navigator.userAgent")),
            "viewportHeight": iDefaultViewportHeight,
            "viewportWidth": iDefaultViewportWidth,
        }
        lRows = buildCsvRows(dResults, dMetadata)
        pathlib.Path(pathOutputDir, sJsonName).write_text(json.dumps({"metadata": dMetadata, "results": dResults}, indent=2, ensure_ascii=False), encoding="utf-8")
        writeCsv(pathlib.Path(pathOutputDir, sCsvName), lRows)
        print("  Capturing page snapshot and screenshot ...")
        sSnapshot = getPageSnapshot(oPage, sNormalizedUrl)
        pathlib.Path(pathOutputDir, sSourceName).write_text(sSnapshot, encoding="utf-8")
        try:
            oPage.screenshot(path=str(pathlib.Path(pathOutputDir, sScreenshotName)), full_page=bDefaultFullPage, timeout=30000)
        except Exception:
            try:
                oPage.screenshot(path=str(pathlib.Path(pathOutputDir, sScreenshotName)), full_page=False, timeout=15000)
            except Exception:
                pass
        pathlib.Path(pathOutputDir, sReportName).write_text(buildReportHtml(dResults, dMetadata, lRows), encoding="utf-8")
        writeReportWorkbook(pathlib.Path(pathOutputDir, sReportWorkbookName), dResults, dMetadata, lRows)
        print(buildConsoleSummary(dResults, dMetadata, str(pathOutputDir)))
        if bOpenReport: os.startfile(str(pathlib.Path(pathOutputDir, sReportName)))
        return str(pathOutputDir)
    finally:
        try:
            if oPage is not None: oPage.close()
        except Exception:
            pass


def styleWorksheet(worksheet):
    cell = None
    iColumnIndex = 0
    lWidths = []
    sColumnLetter = ""

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    for cell in worksheet["1:1"]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(fill_type="solid", fgColor="1F4E78")
        cell.alignment = Alignment(vertical="top", wrap_text=True)
    lWidths = [18, 24, 20, 16, 18, 18, 18, 18, 16, 16, 20, 28, 28, 28, 18, 14, 40, 60, 60]
    for iColumnIndex, iWidth in enumerate(lWidths, start=1):
        sColumnLetter = get_column_letter(iColumnIndex)
        worksheet.column_dimensions[sColumnLetter].width = iWidth
    return worksheet.title


def writeCsv(pathCsv, lRows):
    dRow = {}
    lFieldNames = [
        "scanTimestampUtc", "pageTitle", "pageUrl", "browserVersion", "axeSource",
        "outcome", "ruleId", "impact", "description", "help", "helpUrl",
        "tags", "wcagRefs", "standardsRefs",
        "instanceCount", "instanceIndex", "path", "snippet", "failureSummary",
    ]
    oCsvFile = None

    with pathCsv.open("w", newline="", encoding="utf-8") as oCsvFile:
        csvWriter = csv.DictWriter(oCsvFile, fieldnames=lFieldNames, quoting=csv.QUOTE_ALL, extrasaction="ignore")
        csvWriter.writeheader()
        for dRow in lRows:
            dOut = dict(dRow)
            dOut["instanceCount"] = dRow.get("ruleNodeCount", "")
            dOut["instanceIndex"] = dRow.get("ruleNodeIndex", "")
            dOut["path"]          = dRow.get("target", "")
            dOut["snippet"]       = dRow.get("html", "")
            csvWriter.writerow(dOut)
    return str(pathCsv)


def writeReportWorkbook(pathWorkbook, dResults, dMetadata, lRows):
    cell = None
    iRow = 0
    iSheetIndex = 0
    lImpactRows = []
    lRuleRows = []
    lRow = []
    lWcagRows = []
    workbook = None
    worksheet = None

    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Metadata"
    worksheet.append(["Field", "Value"])
    for lRow in [
        ["Program", sProgramName],
        ["Version", sProgramVersion],
        ["Input", str(dMetadata.get("inputValue") or "")],
        ["Normalized URL", str(dMetadata.get("normalizedUrl") or "")],
        ["Page URL", str(dMetadata.get("pageUrl") or "")],
        ["Page title", str(dMetadata.get("pageTitle") or "")],
        ["Scan timestamp (UTC)", str(dMetadata.get("scanTimestampUtc") or "")],
        ["Browser channel", str(dMetadata.get("browserChannel") or "")],
        ["Browser version", str(dMetadata.get("browserVersion") or "")],
        ["User agent", str(dMetadata.get("userAgent") or "")],
        ["Viewport width", int(dMetadata.get("viewportWidth") or 0)],
        ["Viewport height", int(dMetadata.get("viewportHeight") or 0)],
        ["Navigation timeout (ms)", int(dMetadata.get("navTimeoutMs") or 0)],
        ["Post-load delay (ms)", int(dMetadata.get("postLoadDelayMs") or 0)],
        ["axe-core source", str(dMetadata.get("axeSource") or "")],
        ["axe result types", ", ".join(aAxeRunOptions.get("resultTypes", []))],
        ["Screenshot file", sScreenshotName],
        ["JSON file", sJsonName],
        ["CSV file", sCsvName],
        ["Source snapshot file", sSourceName],
        ["HTML report file", sReportName],
        ["Workbook file", sReportWorkbookName],
    ]:
        worksheet.append(lRow)
    worksheet.column_dimensions["A"].width = 28
    worksheet.column_dimensions["B"].width = 90
    worksheet["A1"].font = Font(bold=True, color="FFFFFF")
    worksheet["B1"].font = Font(bold=True, color="FFFFFF")
    worksheet["A1"].fill = PatternFill(fill_type="solid", fgColor="1F4E78")
    worksheet["B1"].fill = PatternFill(fill_type="solid", fgColor="1F4E78")
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    for cell in worksheet["B"]: cell.alignment = Alignment(vertical="top", wrap_text=True)

    worksheet = workbook.create_sheet("Summary")
    worksheet.append(["Section", "Name", "Value"])
    lImpactRows = getImpactRows(lRows)
    lRuleRows = getRuleFrequencyRows(lRows)
    lWcagRows = getWcagFrequencyRows(lRows)
    worksheet.append(["Overview", "Page title", str(dMetadata.get("pageTitle") or "")])
    worksheet.append(["Overview", "Page URL", str(dMetadata.get("pageUrl") or "")])
    worksheet.append(["Overview", "Violations (rules)", len(dResults.get("violations", []))])
    worksheet.append(["Overview", "Failed instances", sum(len(dRule.get("nodes", [])) for dRule in dResults.get("violations", []))])
    worksheet.append(["Overview", "Needs Review (rules)", len(dResults.get("incomplete", []))])
    worksheet.append(["Overview", "Needs Review (instances)", sum(len(dRule.get("nodes", [])) for dRule in dResults.get("incomplete", []))])
    worksheet.append(["Overview", "Passes (rules)", len(dResults.get("passes", []))])
    worksheet.append(["Overview", "Inapplicable (rules)", len(dResults.get("inapplicable", []))])
    for lRow in lImpactRows: worksheet.append(["Impact (failed instances)", str(lRow[0]), int(lRow[1])])
    for lRow in lRuleRows: worksheet.append(["Top rules by instance count", str(lRow[0]), int(lRow[1])])
    for lRow in lWcagRows:
        sRef = str(lRow[0])
        tInfo = getWcagScInfo(sRef)
        sLabel = f"{sRef} — {tInfo[0]} (Level {tInfo[1]})" if tInfo else sRef
        worksheet.append(["Top WCAG criteria", sLabel, int(lRow[1])])
    worksheet.column_dimensions["A"].width = 28
    worksheet.column_dimensions["B"].width = 56
    worksheet.column_dimensions["C"].width = 16
    styleWorksheet(worksheet)

    worksheet = workbook.create_sheet("Results")
    worksheet.append(["Scan timestamp (UTC)", "Page title", "Page URL", "Browser version", "axe-core source", "Outcome", "Rule ID", "Impact", "Description", "Help", "Help URL", "Tags", "WCAG criteria", "Standards refs", "Instance count", "Instance index", "Path (CSS selector)", "Snippet (HTML)", "Failure summary"])
    for lRow in lRows:
        worksheet.append([lRow["scanTimestampUtc"], lRow["pageTitle"], lRow["pageUrl"], lRow["browserVersion"], lRow["axeSource"], lRow["outcome"], lRow["ruleId"], lRow["impact"], lRow["description"], lRow["help"], lRow["helpUrl"], lRow["tags"], lRow["wcagRefs"], lRow["standardsRefs"], lRow["ruleNodeCount"], lRow["ruleNodeIndex"], lRow["target"], lRow["html"], lRow["failureSummary"]])
    styleWorksheet(worksheet)

    worksheet = workbook.create_sheet("Glossary")
    worksheet.append(["Term", "Definition"])
    for lRow in aGlossaryRows: worksheet.append(lRow)
    worksheet.append(["", ""])
    worksheet.append(["Procedure steps", ""])
    for sProcedure in aProcedures: worksheet.append(["Procedure", sProcedure])
    worksheet.column_dimensions["A"].width = 24
    worksheet.column_dimensions["B"].width = 100
    worksheet["A1"].font = Font(bold=True, color="FFFFFF")
    worksheet["B1"].font = Font(bold=True, color="FFFFFF")
    worksheet["A1"].fill = PatternFill(fill_type="solid", fgColor="1F4E78")
    worksheet["B1"].fill = PatternFill(fill_type="solid", fgColor="1F4E78")
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    for cell in worksheet["B"]: cell.alignment = Alignment(vertical="top", wrap_text=True)

    for iSheetIndex in range(len(workbook.sheetnames)):
        worksheet = workbook[workbook.sheetnames[iSheetIndex]]
        for iRow in range(2, worksheet.max_row + 1): worksheet.row_dimensions[iRow].height = 30
    workbook.save(pathWorkbook)
    return str(pathWorkbook)


# --- Entry point ---

def main():
    bOpenReport = True
    iErrorCount = 0
    iUrlIndex = 0
    iUrlTotal = 0
    lUrls = []
    sInput = ""
    sNormalizedUrl = ""

    arguments = None
    browser = None
    browserType = None
    context = None
    lArgs = []
    pathBaseDir = pathlib.Path.cwd()
    pathOutputDir = None
    playwrightCtx = None

    # Playwright suppresses the default SIGINT handler. Restore it so Ctrl+C works.
    signal.signal(signal.SIGINT, signal.SIG_DFL)

    print(f"{sProgramName} {sProgramVersion}")

    arguments = parseArguments()
    if not arguments.inputValue:
        print(sUsage)
        return 1

    sInput = str(arguments.inputValue)

    # Determine whether sInput is a URL list file or a single URL / local HTML file.
    if isUrlListFile(sInput):
        try:
            lUrls = getUrlsFromFile(sInput)
        except Exception as oError:
            print(f"Error reading URL list file: {oError}", file=sys.stderr)
            return 1
        bOpenReport = False
        iUrlTotal = len(lUrls)
        print(f"URL list: {sInput} ({iUrlTotal} URL(s))")
    else:
        sNormalizedUrl = getNormalizedUrl(sInput)
        lUrls = [sNormalizedUrl]
        iUrlTotal = 1

    try:
        with sync_playwright() as playwrightCtx:
            browserType = playwrightCtx.chromium
            lArgs = [
                "--mute-audio",
                "--no-default-browser-check",
                "--no-first-run",
                f"--window-size={iDefaultViewportWidth},{iDefaultViewportHeight}",
            ]
            browser = browserType.launch(channel=sBrowserChannel, headless=bDefaultHeadless, args=lArgs)
            context = browser.new_context(ignore_https_errors=bDefaultIgnoreHttpsErrors, user_agent=sUserAgent, viewport={"width": iDefaultViewportWidth, "height": iDefaultViewportHeight})
            iUrlIndex = 0
            for sUrl in lUrls:
                iUrlIndex += 1
                sNormalizedUrl = getNormalizedUrl(sUrl) if isUrlListFile(sInput) else sUrl
                if iUrlTotal > 1: print(f"\n[{iUrlIndex}/{iUrlTotal}] {sNormalizedUrl}")
                try:
                    scanUrl(sUrl if isUrlListFile(sInput) else sInput, sNormalizedUrl, browser, context, pathBaseDir, bOpenReport, arguments.wait * 1000)
                except Exception as oError:
                    iErrorCount += 1
                    print(f"  Error scanning {sNormalizedUrl}: {oError}", file=sys.stderr)
                    pathOutputDir = chooseOutputDir(pathBaseDir, sFallbackTitle)
                    pathlib.Path(pathOutputDir, sErrorReportName).write_text("\n\n".join([str(oError), traceback.format_exc()[:iMaxErrorTextLen]]), encoding="utf-8")
            if iUrlTotal > 1:
                print(f"\nDone. {iUrlTotal - iErrorCount} of {iUrlTotal} URLs scanned successfully.")
                if iErrorCount > 0: print(f"  {iErrorCount} error(s). See error.txt in the relevant output directories.")
    finally:
        try:
            if context is not None: context.close()
        except Exception:
            pass
        try:
            if browser is not None: browser.close()
        except Exception:
            pass

    return 0 if iErrorCount == 0 else 1


if __name__ == "__main__": raise SystemExit(main())

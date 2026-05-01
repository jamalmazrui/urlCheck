import argparse, csv, ctypes, datetime, html, io, json, os, pathlib, platform, re, shutil, signal, struct, subprocess, sys, traceback, urllib.error, urllib.parse, urllib.request

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from playwright.sync_api import sync_playwright

# pythonnet (the `clr` module that bridges into the .NET Framework) is imported
# lazily inside showGuiDialog and showFinalGuiMessage so the CLI path has no
# GUI cost or hard runtime dependency on it. The .NET Framework 4.8 ships
# with every supported Windows 10 (since 1903) and Windows 11; no extra
# runtime install is required.


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
iLayoutButtonHeight = 26
iLayoutButtonWidth = 130
iLayoutFormWidth = 600
iLayoutGap = 7
iLayoutLabelWidth = 130
iLayoutLeft = 12
iLayoutRight = 12
iLayoutRowGap = 11
iLayoutTextHeight = 23
iLayoutTop = 12
iMaxTitleLen = 80
iNetworkIdleTimeoutMs = 8000

sAccessibilityInsightsUrl = "https://accessibilityinsights.io/docs/web/overview/"
sAccessibilityYamlName = "page.yaml"
sBrowserChannel = "msedge"
sConfigDirName = "urlCheck"
sConfigFileName = "urlCheck.ini"
sCsvName = "report.csv"
sFallbackTitle = "untitled-page"
sJsonName = "results.json"
sLogFileName = "urlCheck.log"
sMsAccessibilityUrl = "https://learn.microsoft.com/accessibility/"
sProgramName = "urlCheck"
sProgramVersion = "1.10.0"
sReportName = "report.htm"
sReportWorkbookName = "report.xlsx"
sScreenshotName = "page.png"
sSourceName = "page.htm"
sUsage = "Usage: urlCheck [options] <url, domain, local html file, or url-list text file>"
sUserAgent = "urlCheck/1.10.0 (+Playwright Python + axe-core)"
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
        f"<li><a href=\"{html.escape(sAccessibilityYamlName)}\">{html.escape(sAccessibilityYamlName)}</a> — ARIA accessibility tree</li>",
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


def cleanPreviousTempDirs():
    """Remove _MEI* temporary directories in %TEMP% left by previous runs of
    this program. The current run's own directory (sys._MEIPASS) is skipped.
    Directories belonging to currently running PyInstaller applications cannot
    be deleted because their DLLs are locked in memory; shutil.rmtree will
    raise an exception on the first locked file and the whole directory is left
    intact. Only fully-exited runs leave unlocked directories, so this is safe
    to run against all _MEI* siblings regardless of which program created them.
    Has no effect when running from source (sys._MEIPASS is absent)."""
    import shutil
    pathDir = None
    pathTemp = None
    sCurrentMei = ""

    sCurrentMei = getattr(sys, "_MEIPASS", "")
    if not sCurrentMei: return
    pathTemp = pathlib.Path(sCurrentMei).parent
    for pathDir in pathTemp.glob("_MEI*"):
        if pathDir == pathlib.Path(sCurrentMei): continue
        try:
            shutil.rmtree(str(pathDir))
        except Exception:
            pass


def chooseOutputDir(pathBaseDir, sPageTitle, bForce=False):
    """Decide the per-page output folder.

    Returns the Path of the folder to write into, OR None if the folder
    already exists and bForce is False (caller should skip this URL).

    When the folder does not exist, it is created.
    When the folder exists and bForce is True, its contents are deleted
    so the run starts from a clean slate; the folder itself is reused.
    """
    pathCandidate = None
    sBaseName = ""

    sBaseName = getSafeTitle(sPageTitle)
    pathCandidate = pathBaseDir / sBaseName
    if not pathCandidate.exists():
        pathCandidate.mkdir(parents=True, exist_ok=False)
        return pathCandidate
    if bForce:
        # Existing folder, --force is on: empty its contents and reuse
        # the folder. Delete files and subdirectories defensively; if
        # any single entry can't be removed (e.g. a file is open in
        # another process), log it and continue -- the new run will
        # overwrite what it can.
        for child in pathCandidate.iterdir():
            try:
                if child.is_dir() and not child.is_symlink():
                    shutil.rmtree(child)
                else:
                    child.unlink()
            except Exception as ex:
                logger.info(f"Could not remove {child} while emptying "
                    f"{pathCandidate} for --force: {ex}")
        return pathCandidate
    # Existing folder, no --force: caller should skip.
    return None


def ensureSuccess(bCondition, sMessage):
    if not bCondition: raise RuntimeError(sMessage)
    return True


def fetchText(sUrl):
    dHeaders = {}
    request = None
    response = None
    sText = ""

    dHeaders = {"User-Agent": sUserAgent}
    request = urllib.request.Request(sUrl, headers=dHeaders)
    with urllib.request.urlopen(request, timeout=iCdnTimeoutSec) as response: sText = response.read().decode("utf-8")
    return sText


def getAxeScript(page, sPreFetchedContent=""):
    ex = None
    sAxeScript = ""
    sUrl = ""

    # If we have pre-fetched content, inject it directly — this bypasses CDN
    # reachability issues and avoids CSP blocks on external script URLs.
    if sPreFetchedContent:
        for sUrl in aAxeCdnUrls:
            try:
                page.add_script_tag(content=sPreFetchedContent)
                return sUrl
            except Exception as ex:
                break
    # Fall back to URL injection then content fetch per CDN URL
    for sUrl in aAxeCdnUrls:
        try:
            page.add_script_tag(url=sUrl)
            return sUrl
        except Exception as ex:
            continue
    for sUrl in aAxeCdnUrls:
        try:
            sAxeScript = fetchText(sUrl)
            page.add_script_tag(content=sAxeScript)
            return sUrl
        except Exception as ex:
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
    scheme, and local files. Any existing local file is treated as HTML to be
    loaded in the browser, regardless of its extension; the user is
    responsible for supplying an HTML-renderable file. Unrecognised inputs
    are returned unchanged."""
    path = None
    sCandidate = ""
    sSuffix = ""
    sLower = ""

    sCandidate = sInput.strip().strip('"')
    sLower = sCandidate.lower()
    # Already has a scheme — pass through as-is
    if "://" in sLower: return sCandidate
    # Existing local file -- treat as HTML regardless of extension. urlCheck's
    # contract leaves it to the user to ensure the file is HTML-renderable.
    try:
        path = pathlib.Path(sCandidate).expanduser().resolve()
        if path.exists() and path.is_file():
            return path.as_uri()
    except Exception:
        pass
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


def getInitialBrowseDir(sFieldText):
    """Return a directory to use as the initial location of a file or folder
    picker.

    The strategy follows Microsoft's guidance: start at the user's most
    recent / current choice when one exists, otherwise fall back to the
    user's Documents folder.

    sFieldText is the current text of the source-input or output-directory
    field. The function:

      - returns Documents if sFieldText is empty or looks like a URL or
        domain (urlCheck source field can contain those, which are not
        filesystem paths)
      - looks at the first space-separated token if sFieldText has many
      - returns sFieldText itself if it points to an existing directory
      - returns the parent of sFieldText if it points to an existing file
      - returns the parent of sFieldText if only the parent exists
      - falls back to Documents otherwise

    Returns a string. Always returns a non-empty path that the OS knows.
    """
    sCandidate = ""
    sFirstToken = ""
    sParent = ""

    sFieldText = (sFieldText or "").strip()
    if not sFieldText:
        return getDocumentsDir()
    # Source field may contain space-separated tokens. Inspect the first.
    sFirstToken = sFieldText.split()[0] if sFieldText.split() else ""
    if not sFirstToken:
        return getDocumentsDir()
    # Strip surrounding quotes a user may have typed.
    if len(sFirstToken) >= 2 and sFirstToken[0] == '"' and sFirstToken[-1] == '"':
        sFirstToken = sFirstToken[1:-1]
    # URL or domain -- not a filesystem location. Use Documents.
    if re.match(r"^[a-z][a-z0-9+.-]*://", sFirstToken, re.IGNORECASE):
        return getDocumentsDir()
    if re.match(r"^[a-z0-9-]+(\.[a-z0-9-]+)+$", sFirstToken, re.IGNORECASE) and "/" not in sFirstToken and "\\" not in sFirstToken:
        # Bare-domain heuristic: contains a dot, no slashes. Treat as URL.
        return getDocumentsDir()
    # Try as a path. If it has wildcards, strip the basename and inspect
    # the parent directory.
    sCandidate = sFirstToken
    try:
        if any(c in sCandidate for c in "*?["):
            sCandidate = os.path.dirname(sCandidate)
        if sCandidate and os.path.isdir(sCandidate):
            return os.path.abspath(sCandidate)
        if sCandidate and os.path.isfile(sCandidate):
            return os.path.dirname(os.path.abspath(sCandidate))
        sParent = os.path.dirname(sCandidate) if sCandidate else ""
        if sParent and os.path.isdir(sParent):
            return os.path.abspath(sParent)
    except Exception:
        pass
    return getDocumentsDir()


def getDocumentsDir():
    """Return the user's Documents folder path as a string.

    On Windows, this resolves to the per-user Documents folder via
    Environment.SpecialFolder.MyDocuments (the SHGetKnownFolderPath
    KNOWNFOLDERID_Documents). On other platforms (developer machines
    only, urlCheck is Windows-only at runtime), falls back to the
    user's home directory.

    Always returns a non-empty string. If the Documents folder cannot be
    resolved for any reason, returns the home directory.
    """
    sPath = ""
    try:
        # The standard pythonic way: os.path.expanduser handles Windows
        # via the USERPROFILE environment variable, then we append
        # "Documents". But the more correct way on Windows is to query
        # the shell folder, since the user may have redirected
        # Documents to OneDrive or another non-default location.
        # We try SHGetFolderPathW first; if that fails, fall back to
        # %USERPROFILE%\Documents.
        if sys.platform == "win32":
            try:
                bufPath = ctypes.create_unicode_buffer(260)
                # CSIDL_PERSONAL = 0x0005 (Documents). Flags = 0 (current
                # location, not default). HResult 0 (S_OK) on success.
                iCsidlPersonal = 0x0005
                iHr = ctypes.windll.shell32.SHGetFolderPathW(
                    None, iCsidlPersonal, None, 0, bufPath)
                if iHr == 0 and bufPath.value and os.path.isdir(bufPath.value):
                    return bufPath.value
            except Exception:
                pass
        # Cross-platform fallback.
        sPath = os.path.join(os.path.expanduser("~"), "Documents")
        if os.path.isdir(sPath):
            return sPath
        return os.path.expanduser("~")
    except Exception:
        return os.path.expanduser("~")


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
    """Read a plain text file and return a list of non-blank URLs/paths.

    Each line is stripped; lines that are blank or start with # are ignored.
    The file's contents are sniffed up-front: if the leading bytes look
    binary (NUL bytes or a high ratio of non-printable bytes), we raise a
    clear error rather than producing garbage URLs. Lines that don't look
    URL-like are also rejected with a precise line-number error message.

    Raises ValueError on parse failure with a message suitable for the
    user. Raises OSError (FileNotFoundError, PermissionError) if the file
    cannot be opened.
    """
    bLikelyBinary = False
    bytesHead = b""
    iBlankLines = 0
    iLineNo = 0
    iSampleLen = 4096
    lUrls = []
    path = None
    sLine = ""
    sNormalized = ""

    path = pathlib.Path(sInput).expanduser().resolve()

    # Sniff the first few KB to catch binary files. NUL bytes are a strong
    # signal; otherwise any chunk where >30% of the bytes are outside the
    # printable-ASCII / common-control range is treated as binary.
    try:
        with path.open("rb") as fileRaw:
            bytesHead = fileRaw.read(iSampleLen)
    except Exception as ex:
        raise OSError(f"Could not read {sInput}: {ex}")
    if b"\x00" in bytesHead:
        bLikelyBinary = True
    elif bytesHead:
        iPrintable = sum(1 for b in bytesHead
            if (32 <= b < 127) or b in (9, 10, 13))
        if iPrintable / len(bytesHead) < 0.7:
            bLikelyBinary = True
    if bLikelyBinary:
        raise ValueError(
            f"File does not appear to be plain text: {sInput}\n"
            "urlCheck expects a plain text file with one URL per line. "
            "If this is a Word, Excel, PDF, or other binary file, export "
            "or save it as plain text first.")

    # Now read as text. Use UTF-8 with strict error handling to surface
    # encoding problems explicitly; UTF-8-with-BOM is also accepted.
    try:
        with path.open("r", encoding="utf-8-sig", errors="strict") as file:
            iLineNo = 0
            for sLine in file:
                iLineNo += 1
                sLine = sLine.strip()
                if not sLine:
                    iBlankLines += 1
                    continue
                if sLine.startswith("#"): continue
                # Cheap sanity check: each non-comment line should look
                # like a URL, a domain, or a local file path.
                if not _looksLikeUrlOrPath(sLine):
                    raise ValueError(
                        f"{sInput}: line {iLineNo} does not look like a "
                        f"URL, domain, or file path: {sLine!r}\n"
                        "Each non-blank, non-comment line should be one "
                        "URL, one domain name, or one local HTML file path.")
                lUrls.append(sLine)
    except UnicodeDecodeError as ex:
        raise ValueError(
            f"File is not valid UTF-8 text: {sInput}\n"
            f"Decode error: {ex}\n"
            "Re-save the file as plain UTF-8 text and try again.")
    if not lUrls:
        raise ValueError(f"No URLs found in file: {sInput}")
    return lUrls


def _looksLikeUrlOrPath(sLine):
    """Heuristic: does sLine look like a URL, a bare domain, or a file path?

    Used by getUrlsFromFile to catch obviously-wrong lines (e.g. random
    English text from a misclassified document). Generous on purpose --
    we'd rather pass through one bad line and have Playwright reject it
    with a clear error than refuse a legitimate URL because we got the
    pattern wrong. Rejects only lines that contain whitespace or that
    are pure ASCII text with no dot, slash, or colon.
    """
    # Lines with embedded whitespace are never valid (URLs and paths
    # don't contain bare whitespace; if they did, the user should
    # quote/encode).
    if any(c.isspace() for c in sLine): return False
    # A URL has a colon (after the scheme), a path has slashes or a drive
    # letter, a domain has at least one dot. Anything with at least one of
    # these structural characters is plausible.
    return ("://" in sLine) or ("." in sLine) or ("/" in sLine) or ("\\" in sLine)


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
    """Return True if sInput is a path to an existing file (any extension).

    The user is responsible for supplying a plain text file. If a file is
    given that turns out to be binary or non-URL-like, getUrlsFromFile will
    fail with a clear error.
    """
    path = None
    if not sInput: return False
    try:
        path = pathlib.Path(sInput).expanduser()
    except Exception:
        return False
    try:
        if not path.is_file(): return False
    except Exception:
        return False
    return True


def classifyInput(sInput):
    """Classify the source input string.

    Returns one of:
      ('listfile', sPath)  -- sInput is a path to an existing file (any
                              extension), to be parsed as plain text with
                              one URL or local file path per line.
      ('urls', lUrls)      -- sInput is one or more space-separated URLs/
                              domains. lUrls is the list of tokens.
      ('error', sReason)   -- sInput is invalid (currently only the empty
                              case; bad file contents are surfaced later
                              by getUrlsFromFile with a precise message).
    """
    sStripped = ""
    sStripped = (sInput or "").strip()
    if not sStripped:
        return ("error", "No input provided.")
    if isUrlListFile(sStripped):
        return ("listfile", sStripped)
    # Not a file path. Treat as one or more space-separated URLs/domains.
    lTokens = sStripped.split()
    return ("urls", lTokens)


def parseArguments():
    argParser = None

    argParser = argparse.ArgumentParser(
        prog=sProgramName,
        description=(
            "Check one or more web pages for accessibility problems and save "
            "a set of output files in a folder named after each page title. "
            "Pass URLs as separate arguments, or pass the path to a single "
            "plain text file that lists URLs, domains, or local file paths -- "
            "one per line. The list file may have any extension; urlCheck "
            "verifies it is plain text by inspecting its contents."
        ),
        epilog=(
            f"Single URL:    {sProgramName} https://example.com\n"
            f"Domain only:   {sProgramName} microsoft.com\n"
            f"Several URLs:  {sProgramName} https://a.com https://b.com https://c.com\n"
            f"URL list file: {sProgramName} urls.txt\n"
            f"GUI dialog:    {sProgramName} -g\n"
            f"\n"
            f"Output files (in a folder named after each page title):\n"
            f"  report.htm   Accessibility report with headings and links\n"
            f"  report.csv   Violations as a spreadsheet, one row per issue\n"
            f"  report.xlsx  Excel workbook with summary and full results\n"
            f"  results.json Full raw scan data including all metadata\n"
            f"  page.yaml    ARIA accessibility tree of the page\n"
            f"  page.htm     Saved page source with styles inlined\n"
            f"  page.png     Full-page screenshot"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    argParser.add_argument(
        "sSource",
        nargs="*",
        help=(
            "One or more URLs (or domain names) separated by spaces, or the "
            "path to a single plain text file that lists URLs, domains, or "
            "local file paths one per line. The list file may have any "
            "extension. Files referenced inside the list are loaded as HTML "
            "in the browser regardless of extension. Blank lines and lines "
            "starting with # are ignored."
        ),
    )
    argParser.add_argument("-v", "--version", action="version", version=f"%(prog)s {sProgramVersion}")
    argParser.add_argument("-g", "--gui-mode", dest="bGuiMode", action="store_true",
        help="Show the parameter dialog. GUI mode is also entered automatically when urlCheck is launched without arguments from a GUI shell (File Explorer, Start menu, desktop hotkey).")
    argParser.add_argument("-o", "--output-dir", dest="sOutputDir", default="",
        help="Parent directory under which the per-scan output folder is created. Defaults to the current working directory. Created if it does not exist. The per-scan folder is always uniquely named based on the page title.")
    argParser.add_argument("--view-output", dest="bViewOutput", action="store_true",
        help="After all scans complete, open the parent output directory (the -o directory, or the current working directory) in File Explorer.")
    argParser.add_argument("-u", "--use-configuration", dest="bUseConfig", action="store_true",
        help="Load saved settings from %%LOCALAPPDATA%%\\urlCheck\\urlCheck.ini at startup, and write them back on OK in GUI mode. Without this flag urlCheck leaves no filesystem footprint of its own.")
    argParser.add_argument("-l", "--log", dest="bLog", action="store_true",
        help="Write detailed diagnostics to urlCheck.log in the current working directory (UTF-8 with BOM). Any prior urlCheck.log is deleted at the start of the run, so the file always reflects only the current session.")
    argParser.add_argument("-f", "--force", dest="bForce", action="store_true",
        help="Reuse an existing per-page output folder by emptying its contents and writing a fresh set of files. Without this flag urlCheck skips a URL whose per-page output folder already exists, so previous scans are preserved.")
    argParser.add_argument("-i", "--invisible", dest="bInvisible", action="store_true",
        help="Run Microsoft Edge invisibly (the headless browser mode): no visible browser window during the scan.")
    return argParser.parse_args()


def scanUrl(sInput, sNormalizedUrl, browser, context, pathBaseDir, sAxeContent="", bForce=False):
    """Run a single-URL scan.

    Returns:
      - the output directory path string on a successful scan
      - the sentinel string "skipped" if the per-page folder already
        exists and bForce is False (caller decided not to overwrite
        previous results)

    Raises on real errors (navigation failure, axe failure, etc.).
    """
    dMetadata = {}
    dResults = {}
    lArgs = []
    lRows = []
    page = None
    pathOutputDir = None
    sAxeSource = ""
    sPageTitle = ""
    sResultsJson = ""
    sSnapshot = ""

    try:
        page = context.new_page()
        logger.info(f"Navigating to {sNormalizedUrl}")
        page.goto(sNormalizedUrl, timeout=iDefaultNavTimeoutMs, wait_until="load")
        # After the load event, wait briefly for network to settle.
        # A short networkidle timeout lets SPA content finish rendering on most sites
        # while sites with persistent connections (e.g. WebSocket-heavy SPAs) simply
        # time out and continue after iNetworkIdleTimeoutMs milliseconds.
        try:
            page.wait_for_load_state("networkidle", timeout=iNetworkIdleTimeoutMs)
        except Exception:
            pass
        page.wait_for_timeout(iDefaultPostLoadDelayMs)
        # Scroll to bottom and back to trigger lazy-loaded content, then wait briefly
        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        page.wait_for_timeout(500)
        page.evaluate("window.scrollTo(0, 0)")
        page.wait_for_timeout(500)
        sPageTitle = str(page.title() or sFallbackTitle)
        # Decide the output folder. If chooseOutputDir returns None,
        # the per-page folder already exists and --force is not set,
        # so we skip this URL. Skip happens BEFORE the expensive
        # axe-core run, screenshot, snapshot, and report-writing.
        pathOutputDir = chooseOutputDir(pathBaseDir, sPageTitle, bForce=bForce)
        if pathOutputDir is None:
            sExistingDir = str(pathBaseDir / getSafeTitle(sPageTitle))
            print(f"{sNormalizedUrl}")
            print(f"  Skipping ({pathlib.Path(sExistingDir).name} exists, "
                f"use -f to overwrite)")
            logger.info(f"Skipped (output folder exists, no --force): "
                f"{sNormalizedUrl} -> {sExistingDir}")
            return "skipped"
        logger.info(f"Output directory: {pathOutputDir}")
        logger.info("Running axe-core")
        sAxeSource = getAxeScript(page, sAxeContent)
        ensureSuccess(bool(page.evaluate("() => Boolean(window.axe && window.axe.run)")), "axe-core did not load into the page.")
        sResultsJson = page.evaluate("async (opts) => JSON.stringify(await window.axe.run(document, opts))", aAxeRunOptions)
        dResults = json.loads(sResultsJson)
        dMetadata = {
            "axeSource": sAxeSource,
            "browserChannel": sBrowserChannel,
            "browserVersion": str(browser.version),
            "inputValue": sInput,
            "navTimeoutMs": iDefaultNavTimeoutMs,
            "normalizedUrl": sNormalizedUrl,
            "pageTitle": sPageTitle,
            "pageUrl": str(page.url or sNormalizedUrl),
            "postLoadDelayMs": iDefaultPostLoadDelayMs,
            "programName": sProgramName,
            "programVersion": sProgramVersion,
            "scanTimestampUtc": datetime.datetime.now(datetime.timezone.utc).replace(microsecond=0).isoformat(),
            "userAgent": str(page.evaluate("() => navigator.userAgent")),
            "viewportHeight": iDefaultViewportHeight,
            "viewportWidth": iDefaultViewportWidth,
        }
        lRows = buildCsvRows(dResults, dMetadata)
        pathlib.Path(pathOutputDir, sJsonName).write_text(json.dumps({"metadata": dMetadata, "results": dResults}, indent=2, ensure_ascii=False), encoding="utf-8")
        writeCsv(pathlib.Path(pathOutputDir, sCsvName), lRows)
        logger.info("Capturing page snapshot and screenshot")
        sSnapshot = getPageSnapshot(page, sNormalizedUrl)
        pathlib.Path(pathOutputDir, sSourceName).write_text(sSnapshot, encoding="utf-8")
        try:
            page.screenshot(path=str(pathlib.Path(pathOutputDir, sScreenshotName)), full_page=bDefaultFullPage, timeout=30000)
        except Exception:
            try:
                page.screenshot(path=str(pathlib.Path(pathOutputDir, sScreenshotName)), full_page=False, timeout=15000)
            except Exception:
                pass
        try:
            locatorBody = page.locator("body")
            sYaml = locatorBody.aria_snapshot()
            pathlib.Path(pathOutputDir, sAccessibilityYamlName).write_text(sYaml, encoding="utf-8-sig")
        except Exception:
            pass
        pathlib.Path(pathOutputDir, sReportName).write_text(buildReportHtml(dResults, dMetadata, lRows), encoding="utf-8")
        writeReportWorkbook(pathlib.Path(pathOutputDir, sReportWorkbookName), dResults, dMetadata, lRows)
        # User-visible: just URL and page title for this scan. Detailed
        # violation counts go to report.htm/report.csv/report.xlsx and to
        # the log.
        print(f"{sNormalizedUrl}")
        print(f"  Page title: {sPageTitle}")
        logger.info(buildConsoleSummary(dResults, dMetadata, str(pathOutputDir)))
        return str(pathOutputDir)
    finally:
        try:
            if page is not None: page.close()
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
    fileCsv = None

    with pathCsv.open("w", newline="", encoding="utf-8") as fileCsv:
        csvWriter = csv.DictWriter(fileCsv, fieldnames=lFieldNames, quoting=csv.QUOTE_ALL, extrasaction="ignore")
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
        ["ARIA tree file", sAccessibilityYamlName],
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


# --- GUI launch detection (Windows) ---

def _getParentProcessName():
    """
    Returns the lowercased base filename of this process's parent (e.g.
    'cmd.exe', 'explorer.exe'), or '' if it can't be determined. Uses a
    Toolhelp32 snapshot, which is available on every Windows version this
    program supports.
    """
    if sys.platform != "win32": return ""
    try:
        # We use a couple of structures from kernel32. Define them via ctypes
        # to keep the dependency surface minimal (no pywin32, no psutil).
        import ctypes.wintypes as wt
        iSnapProcess = 0x00000002
        iMaxPath = 260

        class ProcessEntry32(ctypes.Structure):
            _fields_ = [
                ("dwSize", wt.DWORD),
                ("cntUsage", wt.DWORD),
                ("th32ProcessID", wt.DWORD),
                ("th32DefaultHeapID", ctypes.c_void_p),
                ("th32ModuleID", wt.DWORD),
                ("cntThreads", wt.DWORD),
                ("th32ParentProcessID", wt.DWORD),
                ("pcPriClassBase", ctypes.c_long),
                ("dwFlags", wt.DWORD),
                ("szExeFile", ctypes.c_char * iMaxPath),
            ]

        kernel32 = ctypes.windll.kernel32
        kernel32.CreateToolhelp32Snapshot.restype = wt.HANDLE
        kernel32.CreateToolhelp32Snapshot.argtypes = [wt.DWORD, wt.DWORD]
        kernel32.Process32First.argtypes = [wt.HANDLE, ctypes.POINTER(ProcessEntry32)]
        kernel32.Process32Next.argtypes = [wt.HANDLE, ctypes.POINTER(ProcessEntry32)]
        kernel32.CloseHandle.argtypes = [wt.HANDLE]

        iMyPid = int(os.getpid())
        iParentPid = 0
        sParentExe = ""

        hSnap = kernel32.CreateToolhelp32Snapshot(iSnapProcess, 0)
        if not hSnap or hSnap == wt.HANDLE(-1).value: return ""
        try:
            processEntry = ProcessEntry32()
            processEntry.dwSize = ctypes.sizeof(ProcessEntry32)
            # Pass 1: find our own entry to learn the parent PID.
            if not kernel32.Process32First(hSnap, ctypes.byref(processEntry)): return ""
            while True:
                if processEntry.th32ProcessID == iMyPid:
                    iParentPid = int(processEntry.th32ParentProcessID)
                    break
                if not kernel32.Process32Next(hSnap, ctypes.byref(processEntry)): break
            if iParentPid == 0: return ""

            # Pass 2: find the parent entry to learn its exe name. Re-walk
            # because Process32First/Next is unidirectional.
            oEntry2 = ProcessEntry32()
            oEntry2.dwSize = ctypes.sizeof(ProcessEntry32)
            if not kernel32.Process32First(hSnap, ctypes.byref(oEntry2)): return ""
            while True:
                if oEntry2.th32ProcessID == iParentPid:
                    sParentExe = oEntry2.szExeFile.decode("ascii", errors="replace")
                    break
                if not kernel32.Process32Next(hSnap, ctypes.byref(oEntry2)): break
        finally:
            kernel32.CloseHandle(hSnap)

        return os.path.basename(sParentExe).lower()
    except Exception:
        return ""


def isLaunchedFromGui():
    """
    Returns True when this process appears to have been launched by a GUI shell
    rather than from a command-line shell.

    Primary signal: GetConsoleProcessList. The same approach 2htm uses. A
    console-subsystem program inherits its parent shell's console (count >= 2)
    when launched from cmd / PowerShell / Windows Terminal / etc., but is the
    only process on a fresh console (count == 1) when Windows creates a new
    console for a double-clicked exe, a Start-menu shortcut, the Run dialog,
    or a desktop hotkey.

    Secondary signal: parent process name. Used only when the count is
    ambiguous: in particular, count == 0 (truly no console attached, which
    happens in some service / scheduled-task contexts) or when a third party
    has briefly attached a process to our console between launch and the
    moment we make the check (rare but possible -- accessibility tools,
    AV scanners, JAWS / NVDA event hooks, etc.). When count == 1 the answer
    is unambiguous regardless of parent.

    Both signals and the resulting decision are written to the log when -l
    is in effect, so the detection can be diagnosed from the log file.
    """
    bResult = False
    iCount = 0
    iCountErr = 0
    sParent = ""
    sReason = ""

    if sys.platform != "win32":
        sReason = "non-Windows platform; assume CLI"
        bResult = False
    else:
        try:
            aiBuf = (ctypes.c_uint * 16)()
            iCount = int(ctypes.windll.kernel32.GetConsoleProcessList(
                aiBuf, ctypes.c_uint(16)))
        except Exception as ex:
            iCountErr = 1
            iCount = -1
            sReason = f"GetConsoleProcessList failed: {ex}"
        if iCount == 1:
            bResult = True
            sReason = "console process count is 1 (fresh console -> GUI launch)"
        elif iCount >= 2:
            bResult = False
            sReason = f"console process count is {iCount} (sharing parent shell -> CLI)"
        else:
            # iCount == 0 (no console) or iCount < 0 (call failed). Fall
            # back to parent-process inspection.
            sParent = _getParentProcessName()
            sShells = ("cmd.exe", "powershell.exe", "pwsh.exe",
                "windowsterminal.exe", "openconsole.exe", "wt.exe",
                "conemu64.exe", "conemu.exe", "cmder.exe",
                "bash.exe", "wsl.exe", "git-bash.exe", "mintty.exe")
            if sParent in sShells:
                bResult = False
                sReason = f"console count={iCount}, parent={sParent} -> CLI"
            elif sParent:
                bResult = True
                sReason = f"console count={iCount}, parent={sParent} -> GUI"
            else:
                bResult = False
                sReason = f"console count={iCount}, parent unknown -> assume CLI"

    logger.info(f"isLaunchedFromGui: parent='{sParent}' "
        f"consoleCount={iCount} -> bIsGui={bResult} ({sReason})")
    return bResult


def hideOwnConsoleWindow():
    """
    Hides this process's console window. Only safe when isLaunchedFromGui()
    returned True; the console was created by Windows for us in that case and
    no parent shell is sharing it. Non-fatal on failure.
    """
    iSwHide = 0
    if sys.platform != "win32": return
    try:
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd: ctypes.windll.user32.ShowWindow(hwnd, iSwHide)
    except Exception:
        pass


def openFolderInExplorer(sPath):
    """Open the given folder in Windows Explorer. Non-fatal on failure."""
    try:
        if sys.platform == "win32":
            os.startfile(sPath)
        else:
            subprocess.Popen(["xdg-open", sPath])
    except Exception:
        pass


# --- Logger ---

class logger:
    """
    Tiny diagnostic logger written to urlCheck.log in CWD when -l / --log is
    given. UTF-8 with BOM so Notepad opens it correctly. Each session starts
    with a fresh file -- any prior log is deleted before the new one is
    opened, so the log only ever contains output from the current run.
    Open is lazy and silent on failure; a logging error must never sink a
    scan.
    """
    fLog = None
    bEnabled = False

    @classmethod
    def open(cls):
        try:
            # Delete any prior session's log first so the new file contains
            # only this session's output. unlink may fail if the file does
            # not exist (FileNotFoundError) -- that's fine; just continue.
            try:
                os.unlink(sLogFileName)
            except FileNotFoundError:
                pass
            except Exception:
                # If unlink fails for some other reason (e.g., another
                # process has the file open), fall through to open() in
                # "w" mode, which will truncate. open() may also fail in
                # that case, and the outer except will mark logger
                # disabled -- but that's the right outcome: better to lose
                # logging than to mix old and new content.
                pass
            cls.fLog = open(sLogFileName, "w", encoding="utf-8-sig", newline="\n")
            cls.bEnabled = True
            cls.info(f"{sProgramName} {sProgramVersion} log opened (previous urlCheck.log deleted)")
        except Exception:
            cls.bEnabled = False

    @classmethod
    def write(cls, sLevel, sMsg):
        # Internal level-tagged writer. Public level methods (info,
        # warn, error, debug) all funnel through here so the format
        # is uniform and a level filter could be added later in one
        # place. Silent no-op when the logger is disabled, which is
        # the case unless the user passed -l / --log.
        if not cls.bEnabled or cls.fLog is None: return
        try:
            sStamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            cls.fLog.write(f"[{sStamp}] [{sLevel}] {sMsg}\n")
            cls.fLog.flush()
        except Exception:
            pass

    @classmethod
    def info(cls, sMsg):  cls.write("INFO",  sMsg)
    @classmethod
    def warn(cls, sMsg):  cls.write("WARN",  sMsg)
    @classmethod
    def error(cls, sMsg): cls.write("ERROR", sMsg)
    @classmethod
    def debug(cls, sMsg): cls.write("DEBUG", sMsg)

    @classmethod
    def close(cls):
        try:
            if cls.fLog is not None: cls.fLog.close()
        except Exception:
            pass
        cls.fLog = None
        cls.bEnabled = False


# --- Config manager (opt-in, INI under %LOCALAPPDATA%\urlCheck) ---

class configManager:
    """
    Persists user preferences to %LOCALAPPDATA%\\urlCheck\\urlCheck.ini, but
    only when the user opts in via -u / --use-configuration or via the GUI
    checkbox. Without opt-in urlCheck leaves no filesystem footprint of its
    own beyond the per-scan output folders the user explicitly asks for.
    """

    @staticmethod
    def getConfigDir():
        sLocal = os.environ.get("LOCALAPPDATA", "") or os.path.expanduser("~")
        return os.path.join(sLocal, sConfigDirName)

    @staticmethod
    def getConfigPath():
        return os.path.join(configManager.getConfigDir(), sConfigFileName)

    @staticmethod
    def configExists():
        try: return os.path.isfile(configManager.getConfigPath())
        except Exception: return False

    @staticmethod
    def eraseAll():
        sDir = configManager.getConfigDir()
        sPath = configManager.getConfigPath()
        try:
            if os.path.isfile(sPath):
                os.remove(sPath)
                logger.info(f"Deleted configuration file: {sPath}")
        except Exception as ex:
            logger.info(f"Could not delete configuration file {sPath}: {ex}")
        try:
            if os.path.isdir(sDir) and not os.listdir(sDir):
                os.rmdir(sDir)
                logger.info(f"Removed empty configuration directory: {sDir}")
        except Exception as ex:
            logger.info(f"Could not remove configuration directory {sDir}: {ex}")

    @staticmethod
    def parseFile(sPath):
        d = {}
        sLine = ""
        sLineRaw = ""
        with open(sPath, "r", encoding="utf-8-sig") as fIni:
            for sLineRaw in fIni:
                sLine = sLineRaw.strip()
                if not sLine: continue
                if sLine.startswith(";") or sLine.startswith("#"): continue
                if sLine.startswith("[") and sLine.endswith("]"): continue
                iEq = sLine.find("=")
                if iEq <= 0: continue
                d[sLine[:iEq].strip().lower()] = sLine[iEq + 1:].strip()
        return d

    @staticmethod
    def getBool(d, sKey):
        s = (d.get(sKey, "") or "").strip().lower()
        return s in ("1", "true", "yes", "on")

    @staticmethod
    def loadInto(arguments):
        """
        Load saved values into the parsed-arguments object. CLI-supplied
        values always win: only fields the user did NOT specify on the command
        line are overwritten. The cli-supplied set is reconstructed by
        comparing against parser defaults.
        """
        d = {}
        sPath = ""

        sPath = configManager.getConfigPath()
        if not os.path.isfile(sPath): return
        try:
            d = configManager.parseFile(sPath)
        except Exception as ex:
            print(f"[WARN] Could not read configuration from {sPath}: {ex}")
            return

        # Source: only adopt the saved value if the CLI provided no positional.
        if (not getattr(arguments, "sSource", None)) and d.get("source", ""):
            arguments.sSource = d.get("source", "")

        # Output dir: only adopt if the CLI did not pass -o.
        if not getattr(arguments, "sOutputDir", "") and d.get("output_directory", ""):
            arguments.sOutputDir = d.get("output_directory", "")

        # Booleans: only adopt if the CLI did not pass the flag (i.e., it's
        # currently False, the parser default).
        if not getattr(arguments, "bViewOutput", False):
            arguments.bViewOutput = configManager.getBool(d, "view_output")
        if not getattr(arguments, "bInvisible", False):
            arguments.bInvisible = configManager.getBool(d, "invisible")
        if not getattr(arguments, "bForce", False):
            arguments.bForce = configManager.getBool(d, "force_replacements")
        if not getattr(arguments, "bLog", False):
            arguments.bLog = configManager.getBool(d, "log_session")

    @staticmethod
    def save(sSource, sOutputDir, bViewOutput, bInvisible, bForce, bLog):
        sDir = ""
        sPath = ""

        sDir = configManager.getConfigDir()
        sPath = configManager.getConfigPath()
        try:
            if not os.path.isdir(sDir): os.makedirs(sDir, exist_ok=True)
            with open(sPath, "w", encoding="utf-8-sig", newline="\n") as fIni:
                fIni.write("; urlCheck configuration\n")
                fIni.write("; auto-written when Use configuration was checked at OK time.\n")
                fIni.write("; Delete this file to reset, or click Default settings in the\n")
                fIni.write("; GUI, which also deletes the file and the urlCheck folder.\n")
                fIni.write(f"source={sSource or ''}\n")
                fIni.write(f"output_directory={sOutputDir or ''}\n")
                fIni.write(f"view_output={'1' if bViewOutput else '0'}\n")
                fIni.write(f"invisible={'1' if bInvisible else '0'}\n")
                fIni.write(f"force_replacements={'1' if bForce else '0'}\n")
                fIni.write(f"log_session={'1' if bLog else '0'}\n")
            logger.info(f"Saved configuration to {sPath}")
        except Exception as ex:
            print(f"[WARN] Could not save configuration to {sPath}: {ex}")
            logger.info(f"Could not save configuration: {ex}")


# --- GUI dialog (Python.NET / WinForms) ---
#
# Uses pythonnet's `clr` to load the .NET Framework 4.8 System.Windows.Forms
# assembly that ships with every supported version of Windows 10 and 11. This
# gives us the *same* widget set 2htm uses (System.Windows.Forms.Form,
# Label, TextBox, CheckBox, Button, FolderBrowserDialog, OpenFileDialog,
# MessageBox), so screen-reader behavior in JAWS and NVDA is identical
# between the two tools: labels announced via Label-to-control association,
# ampersand access keys, AcceptButton/CancelButton wiring, predictable tab
# order. No extra runtime to install -- the .NET Framework 4.8 is built
# into Windows 10 (since 1903) and Windows 11.
#
# pythonnet (the `clr` module) is imported lazily, so the CLI path does not
# require it to be installed at all. PyInstaller bundles pythonnet's tiny
# loader (clr.pyd, Python.Runtime.dll) when the build script runs; at
# runtime that loader bridges into the in-box .NET Framework.

def _loadDotNetForms():
    """
    Import pythonnet and the System.Windows.Forms / System.Drawing namespaces.
    Returns a dict of the names this file uses, or None on failure (e.g.
    pythonnet not installed). Keeps the import-mess in one place so the
    dialog code below reads cleanly.
    """
    try:
        import clr  # provided by the pythonnet package
        clr.AddReference("System.Windows.Forms")
        clr.AddReference("System.Drawing")
        from System import EventHandler, Environment as DotNetEnvironment
        from System.Drawing import ContentAlignment, Font, FontFamily, Point, Size, SystemFonts
        from System.Threading import ApartmentState, Thread
        from System.Windows.Forms import (
            AnchorStyles, Application, Button, CheckBox, DialogResult,
            DockStyle, FolderBrowserDialog, Form, FormBorderStyle,
            FormStartPosition, Keys, Label, MessageBox, MessageBoxButtons,
            MessageBoxDefaultButton, MessageBoxIcon, OpenFileDialog, Panel,
            ScrollBars, TextBox)
        return {
            "AnchorStyles": AnchorStyles, "ApartmentState": ApartmentState,
            "Application": Application,
            "Button": Button, "CheckBox": CheckBox,
            "ContentAlignment": ContentAlignment, "DialogResult": DialogResult,
            "DockStyle": DockStyle, "DotNetEnvironment": DotNetEnvironment,
            "EventHandler": EventHandler,
            "FolderBrowserDialog": FolderBrowserDialog,
            "Font": Font, "FontFamily": FontFamily, "Form": Form,
            "FormBorderStyle": FormBorderStyle,
            "FormStartPosition": FormStartPosition,
            "Keys": Keys, "Label": Label, "MessageBox": MessageBox,
            "MessageBoxButtons": MessageBoxButtons,
            "MessageBoxDefaultButton": MessageBoxDefaultButton,
            "MessageBoxIcon": MessageBoxIcon, "OpenFileDialog": OpenFileDialog,
            "Panel": Panel, "Point": Point, "ScrollBars": ScrollBars,
            "Size": Size, "SystemFonts": SystemFonts,
            "TextBox": TextBox, "Thread": Thread,
        }
    except Exception as ex:
        print(f"[ERROR] GUI mode requires pythonnet (the `clr` module) and the .NET Framework: {ex}")
        print("        pip install pythonnet")
        return None


def browseForFolderViaShell(sTitle, sInitialPath=""):
    """
    Native folder picker via SHBrowseForFolderW + SHGetPathFromIDListW from
    shell32.dll. Used in place of WinForms' FolderBrowserDialog because the
    latter deadlocks under pythonnet (issue #657).

    We deliberately use the OLDER ("classic") browse-for-folder dialog by
    setting only BIF_RETURNONLYFSDIRS in ulFlags. The newer dialog style
    (BIF_NEWDIALOGSTYLE / BIF_USENEWUI) requires OleInitialize and a
    single-threaded apartment (COINIT_APARTMENTTHREADED); per the Microsoft
    documentation, "If COM is initialized through CoInitializeEx with the
    COINIT_MULTITHREADED flag set, SHBrowseForFolder fails if
    BIF_NEWDIALOGSTYLE is passed." Pythonnet's COM initialization state can
    leave the thread in MTA, so the new dialog style hangs. The classic
    dialog has no such requirement and works reliably.

    The trade-off is cosmetic: the classic dialog has no resize handle, no
    edit field for typing a path, and no "create new folder" button. It
    does have a complete file system tree, which is enough to choose an
    output directory.

    sInitialPath: if non-empty and an existing directory, the dialog opens
    with that folder pre-selected and expanded in the tree. We honor this
    by registering a small BFFCALLBACK that posts BFFM_SETSELECTION on
    BFFM_INITIALIZED. The callback is a plain Win32 function pointer and
    introduces no COM dependencies.

    Returns the chosen folder path as a string, or "" if the user cancelled.
    sTitle is shown above the folder tree.
    """
    import ctypes.wintypes as wt
    iBifReturnOnlyFsDirs = 0x00000001
    iBffmInitialized = 1
    iBffmSetSelection = 0x400 + 103  # WM_USER + 103 == BFFM_SETSELECTIONW
    iMaxPath = 260

    class BrowseInfoW(ctypes.Structure):
        _fields_ = [
            ("hwndOwner", wt.HWND),
            ("pidlRoot", ctypes.c_void_p),
            ("pszDisplayName", wt.LPWSTR),
            ("lpszTitle", wt.LPCWSTR),
            ("ulFlags", wt.UINT),
            ("lpfn", ctypes.c_void_p),
            ("lParam", wt.LPARAM),
            ("iImage", ctypes.c_int),
        ]

    shell32 = ctypes.windll.shell32
    user32 = ctypes.windll.user32
    ole32 = ctypes.windll.ole32

    # Display-name buffer required by the API; we discard it and use
    # SHGetPathFromIDListW for the actual file-system path.
    bufDisplay = ctypes.create_unicode_buffer(iMaxPath)

    # Build the BFFCALLBACK that selects the initial path when the dialog
    # finishes initializing. SendMessage with BFFM_SETSELECTIONW takes a
    # wide-string path in lParam (and TRUE in wParam to indicate it's a
    # path string, not a PIDL).
    user32.SendMessageW.restype = ctypes.c_long
    user32.SendMessageW.argtypes = [wt.HWND, ctypes.c_uint, ctypes.c_void_p, ctypes.c_void_p]

    # Validate and normalize the seed path. If it's empty or doesn't exist,
    # we don't install a callback at all -- the dialog opens at its default
    # (a file-system root view).
    sNormalizedInitial = ""
    if sInitialPath:
        try:
            pathSeed = pathlib.Path(sInitialPath).expanduser().resolve()
            if pathSeed.is_dir():
                sNormalizedInitial = str(pathSeed)
        except Exception:
            sNormalizedInitial = ""

    pCallback = None
    callbackProto = None
    if sNormalizedInitial:
        # Keep the path string alive for the lifetime of the dialog by
        # closing over it in the callback. Without this, ctypes would
        # garbage-collect the wide-string before the dialog reads it.
        sCallbackPath = sNormalizedInitial
        # BFFCALLBACK signature: int CALLBACK Fn(HWND, UINT, LPARAM, LPARAM).
        # We use ctypes.WINFUNCTYPE for the stdcall convention.
        callbackProto = ctypes.WINFUNCTYPE(
            ctypes.c_int, wt.HWND, ctypes.c_uint, wt.LPARAM, wt.LPARAM)
        def fnBrowseCallback(hwnd, iMsg, lParam, lpData):
            if iMsg == iBffmInitialized:
                try:
                    user32.SendMessageW(hwnd, iBffmSetSelection,
                        ctypes.c_void_p(1),  # TRUE: lParam is a string path
                        ctypes.c_wchar_p(sCallbackPath))
                except Exception:
                    pass
            return 0
        pCallback = callbackProto(fnBrowseCallback)

    bi = BrowseInfoW()
    bi.hwndOwner = user32.GetActiveWindow()  # may be NULL; that's OK
    bi.pidlRoot = None
    bi.pszDisplayName = ctypes.cast(bufDisplay, wt.LPWSTR)
    bi.lpszTitle = sTitle
    bi.ulFlags = iBifReturnOnlyFsDirs  # classic dialog only
    bi.lpfn = ctypes.cast(pCallback, ctypes.c_void_p) if pCallback else None
    bi.lParam = 0
    bi.iImage = 0

    # Initialize COM as STA. CoInitialize is shorthand for
    # CoInitializeEx(NULL, COINIT_APARTMENTTHREADED). It returns S_FALSE if
    # the thread already had a compatible apartment; we treat that as fine.
    # If the thread was previously in MTA, this call will return RPC_E_CHANGED_MODE
    # and the classic dialog will still work because it does not require STA.
    shell32.SHBrowseForFolderW.restype = ctypes.c_void_p
    shell32.SHBrowseForFolderW.argtypes = [ctypes.POINTER(BrowseInfoW)]
    shell32.SHGetPathFromIDListW.restype = wt.BOOL
    shell32.SHGetPathFromIDListW.argtypes = [ctypes.c_void_p, wt.LPWSTR]
    ole32.CoTaskMemFree.argtypes = [ctypes.c_void_p]
    try:
        ole32.CoInitialize(None)
    except Exception: pass

    pidl = shell32.SHBrowseForFolderW(ctypes.byref(bi))
    # Keep callbackProto and pCallback alive until SHBrowseForFolderW returns.
    # This line is a no-op at runtime but documents the lifetime requirement
    # so a future reader doesn't accidentally drop the reference.
    _ = (pCallback, callbackProto)
    if not pidl: return ""
    try:
        bufPath = ctypes.create_unicode_buffer(iMaxPath)
        if not shell32.SHGetPathFromIDListW(pidl, bufPath):
            return ""
        return bufPath.value
    finally:
        try: ole32.CoTaskMemFree(pidl)
        except Exception: pass


def showGuiDialog(arguments):
    """
    Show the urlCheck parameter dialog. Mutates `arguments` in place with the
    user's chosen values. Returns True on OK, False on Cancel.

    Layout mirrors 2htm:
      Row 1: "Source URLs:"                  [textbox]   [Browse source]
      Row 2: "Output directory:"             [textbox]   [Choose output]
      Row 3: [x] Invisible mode             [x] View output
      Row 4: [x] Log session                [x] Use configuration
      Row 5: [Help] [Default settings]                  [OK] [Cancel]
    """
    d = _loadDotNetForms()
    logger.info(f"showGuiDialog: _loadDotNetForms returned: {'ok' if d is not None else 'NONE (pythonnet missing?)'}")
    if d is None: return False

    # Pull names out of the namespace dict for readability.
    Application = d["Application"]
    Button = d["Button"]
    CheckBox = d["CheckBox"]
    ContentAlignment = d["ContentAlignment"]
    DialogResult = d["DialogResult"]
    EventHandler = d["EventHandler"]
    FolderBrowserDialog = d["FolderBrowserDialog"]
    Form = d["Form"]
    FormBorderStyle = d["FormBorderStyle"]
    FormStartPosition = d["FormStartPosition"]
    Keys = d["Keys"]
    Label = d["Label"]
    MessageBox = d["MessageBox"]
    MessageBoxButtons = d["MessageBoxButtons"]
    MessageBoxDefaultButton = d["MessageBoxDefaultButton"]
    MessageBoxIcon = d["MessageBoxIcon"]
    OpenFileDialog = d["OpenFileDialog"]
    Point = d["Point"]
    Size = d["Size"]
    SystemFonts = d["SystemFonts"]
    TextBox = d["TextBox"]

    # Log .NET / pythonnet diagnostics so the GUI environment is
    # self-documenting in the log file.
    try:
        environment = d["DotNetEnvironment"]
        logger.info(f".NET runtime: Version={environment.Version} "
            f"OSVersion={environment.OSVersion} Is64Bit={environment.Is64BitProcess}")
    except Exception as ex:
        logger.info(f".NET diagnostics unavailable: {ex}")

    # Set the calling thread to the single-threaded apartment (STA) COM
    # model. WinForms common dialogs -- OpenFileDialog, FolderBrowserDialog,
    # and the various shell extensions they delegate to -- require an STA
    # thread. C# WinForms apps get this for free via [STAThread] on Main();
    # in pythonnet we have to set it explicitly. If this is omitted, the
    # main dialog usually opens fine, but clicking Browse source or Choose
    # output deadlocks the COM marshaler and the program appears to lock up.
    #
    # SetApartmentState fails with InvalidOperationException if the thread
    # has already started COM in MTA mode (e.g. by a prior dialog call in
    # the same process). The wrap-and-log pattern below makes that case
    # diagnosable in the log without crashing.
    try:
        thread = d["Thread"].CurrentThread
        sBefore = str(thread.GetApartmentState())
        thread.SetApartmentState(d["ApartmentState"].STA)
        sAfter = str(thread.GetApartmentState())
        logger.info(f"Thread apartment state: {sBefore} -> {sAfter}")
    except Exception as ex:
        logger.info(f"SetApartmentState(STA) failed (continuing): {ex}")

    # Enable modern Windows visual styles (Common Controls 6 themed widgets:
    # rounded buttons, themed scroll bars, etc.) and the GDI+ TextRenderer
    # for crisper text. Both calls MUST happen before any Form, Button, or
    # other control is constructed in this AppDomain. WinForms ignores them
    # otherwise. EnableVisualStyles is also harmless to call repeatedly, so
    # it is safe even if showGuiDialog is invoked more than once.
    #
    # Accessibility note: visual styles are purely cosmetic. The underlying
    # HWNDs and MSAA / UI Automation properties (control type, name, role,
    # state) are unchanged, so JAWS and NVDA see the exact same accessibility
    # tree they would see with the classic theme. If a regression is
    # observed, comment out these two lines and rebuild.
    try:
        d["Application"].EnableVisualStyles()
        d["Application"].SetCompatibleTextRenderingDefault(False)
    except Exception: pass

    # Initial values from arguments (which may have been pre-loaded from
    # saved config and/or the command line).
    sInitTarget = getattr(arguments, "sSource", "") or ""
    sInitOutDir = getattr(arguments, "sOutputDir", "") or ""
    bInitView = bool(getattr(arguments, "bViewOutput", False))
    bInitInvisible = bool(getattr(arguments, "bInvisible", False))
    bInitForce = bool(getattr(arguments, "bForce", False))
    bInitLog = bool(getattr(arguments, "bLog", False))
    bInitUseCfg = bool(getattr(arguments, "bUseConfig", False))

    # Layout constants live at module scope as iLayout* (see top of file).
    # Match extCheck and 2htm exactly so the three dialogs feel like
    # siblings.
    iFormW = iLayoutFormWidth
    iTextX = iLayoutLeft + iLayoutLabelWidth + iLayoutGap
    iTextW = iFormW - iTextX - iLayoutGap - iLayoutButtonWidth - iLayoutRight
    iBtnX = iFormW - iLayoutRight - iLayoutButtonWidth

    # Build the form.
    frm = Form()
    frm.Text = sProgramName
    frm.FormBorderStyle = FormBorderStyle.FixedDialog
    frm.StartPosition = FormStartPosition.CenterScreen
    frm.MaximizeBox = False
    frm.MinimizeBox = False
    frm.ShowInTaskbar = True
    frm.ClientSize = Size(iFormW, 280)
    frm.Font = SystemFonts.MessageBoxFont

    # F1 -> Help. KeyPreview lets the form see the keystroke before child
    # controls consume it. F1 is the standard Windows help shortcut and is
    # expected by keyboard-driven and screen-reader users.
    frm.KeyPreview = True

    # --- Row 1: Target ---
    y = iLayoutTop
    lblTarget = Label()
    lblTarget.Text = "&Source files:"
    lblTarget.AutoSize = False
    lblTarget.Location = Point(iLayoutLeft, y + 3)
    lblTarget.Size = Size(iLayoutLabelWidth, iLayoutTextHeight)
    lblTarget.TextAlign = ContentAlignment.MiddleLeft
    frm.Controls.Add(lblTarget)

    txtTarget = TextBox()
    txtTarget.Text = sInitTarget
    txtTarget.Location = Point(iTextX, y)
    txtTarget.Size = Size(iTextW, iLayoutTextHeight)
    txtTarget.TabIndex = 0
    frm.Controls.Add(txtTarget)

    btnBrowseTarget = Button()
    btnBrowseTarget.Text = "&Browse source..."
    btnBrowseTarget.Location = Point(iBtnX, y - 1)
    btnBrowseTarget.Size = Size(iLayoutButtonWidth, iLayoutButtonHeight)
    btnBrowseTarget.TabIndex = 1
    btnBrowseTarget.UseVisualStyleBackColor = True
    frm.Controls.Add(btnBrowseTarget)

    # --- Row 2: Output directory ---
    y += iLayoutTextHeight + iLayoutRowGap
    lblOut = Label()
    lblOut.Text = "&Output directory:"
    lblOut.AutoSize = False
    lblOut.Location = Point(iLayoutLeft, y + 3)
    lblOut.Size = Size(iLayoutLabelWidth, iLayoutTextHeight)
    lblOut.TextAlign = ContentAlignment.MiddleLeft
    frm.Controls.Add(lblOut)

    txtOut = TextBox()
    txtOut.Text = sInitOutDir
    txtOut.Location = Point(iTextX, y)
    txtOut.Size = Size(iTextW, iLayoutTextHeight)
    txtOut.TabIndex = 2
    frm.Controls.Add(txtOut)

    btnBrowseOut = Button()
    btnBrowseOut.Text = "&Choose output..."
    btnBrowseOut.Location = Point(iBtnX, y - 1)
    btnBrowseOut.Size = Size(iLayoutButtonWidth, iLayoutButtonHeight)
    btnBrowseOut.TabIndex = 3
    btnBrowseOut.UseVisualStyleBackColor = True
    frm.Controls.Add(btnBrowseOut)

    # --- Row 3: program-specific option row (urlCheck has Invisible) ---
    # The program-specific option appears alone above the common
    # option grid so the layout reads consistently with 2htm and
    # extCheck: program-specific first, common options after.
    y += iLayoutTextHeight + iLayoutRowGap * 2
    iChkW = (iFormW - iLayoutLeft - iLayoutRight) // 2
    chkInvisible = CheckBox()
    chkInvisible.Text = "&Invisible mode"
    chkInvisible.Checked = bInitInvisible
    chkInvisible.Location = Point(iLayoutLeft, y)
    chkInvisible.Size = Size(iChkW, iLayoutTextHeight)
    chkInvisible.TabIndex = 4
    frm.Controls.Add(chkInvisible)

    # --- Row 4: Force replacements + View output (common 2x2 grid, top row) ---
    y += iLayoutTextHeight + iLayoutRowGap
    chkForce = CheckBox()
    chkForce.Text = "&Force replacements"
    chkForce.Checked = bInitForce
    chkForce.Location = Point(iLayoutLeft, y)
    chkForce.Size = Size(iChkW, iLayoutTextHeight)
    chkForce.TabIndex = 5
    frm.Controls.Add(chkForce)

    chkView = CheckBox()
    chkView.Text = "&View output"
    chkView.Checked = bInitView
    chkView.Location = Point(iLayoutLeft + iChkW, y)
    chkView.Size = Size(iChkW, iLayoutTextHeight)
    chkView.TabIndex = 6
    frm.Controls.Add(chkView)

    # --- Row 5: Log session + Use configuration (common 2x2 grid, bottom row) ---
    y += iLayoutTextHeight + iLayoutRowGap
    chkLog = CheckBox()
    chkLog.Text = "&Log session"
    chkLog.Checked = bInitLog
    chkLog.Location = Point(iLayoutLeft, y)
    chkLog.Size = Size(iChkW, iLayoutTextHeight)
    chkLog.TabIndex = 7
    frm.Controls.Add(chkLog)

    chkUseCfg = CheckBox()
    chkUseCfg.Text = "&Use configuration"
    chkUseCfg.Checked = bInitUseCfg
    chkUseCfg.Location = Point(iLayoutLeft + iChkW, y)
    chkUseCfg.Size = Size(iChkW, iLayoutTextHeight)
    chkUseCfg.TabIndex = 8
    frm.Controls.Add(chkUseCfg)

    # --- Bottom row: Help, Defaults on the left; OK, Cancel on the right ---
    y += iLayoutTextHeight + iLayoutRowGap * 2
    btnHelp = Button()
    btnHelp.Text = "&Help"
    btnHelp.Location = Point(iLayoutLeft, y)
    btnHelp.Size = Size(iLayoutButtonWidth, iLayoutButtonHeight)
    btnHelp.TabIndex = 9
    btnHelp.UseVisualStyleBackColor = True
    frm.Controls.Add(btnHelp)

    btnDefaults = Button()
    btnDefaults.Text = "&Default settings"
    btnDefaults.Location = Point(iLayoutLeft + iLayoutButtonWidth + iLayoutGap, y)
    btnDefaults.Size = Size(iLayoutButtonWidth, iLayoutButtonHeight)
    btnDefaults.TabIndex = 10
    btnDefaults.UseVisualStyleBackColor = True
    frm.Controls.Add(btnDefaults)

    btnOk = Button()
    btnOk.Text = "OK"
    btnOk.DialogResult = DialogResult.OK
    btnOk.Location = Point(iFormW - iLayoutRight - 2 * iLayoutButtonWidth - iLayoutGap, y)
    btnOk.Size = Size(iLayoutButtonWidth, iLayoutButtonHeight)
    btnOk.TabIndex = 11
    btnOk.UseVisualStyleBackColor = True
    frm.Controls.Add(btnOk)

    btnCancel = Button()
    btnCancel.Text = "Cancel"
    btnCancel.DialogResult = DialogResult.Cancel
    btnCancel.Location = Point(iBtnX, y)
    btnCancel.Size = Size(iLayoutButtonWidth, iLayoutButtonHeight)
    btnCancel.TabIndex = 12
    btnCancel.UseVisualStyleBackColor = True
    frm.Controls.Add(btnCancel)

    # Wire defaults: Enter -> OK, Esc -> Cancel.
    frm.AcceptButton = btnOk
    frm.CancelButton = btnCancel

    # Adjust form height to accommodate the last control plus a margin.
    frm.ClientSize = Size(iFormW, y + iLayoutButtonHeight + iLayoutTop)

    # --- Event handlers (defined as nested functions so they close over the
    #     control variables above) ---

    def fnShowHelp():
        sMsg = (
            f"{sProgramName} {sProgramVersion} checks one or more web pages "
            f"for accessibility problems and saves a set of output files in "
            f"a folder named after each page title.\r\n\r\n"
            f"Source files: enter one URL (https://example.com), or a domain "
            f"(microsoft.com), or several of either separated by spaces, or "
            f"the path to a single plain text file that lists URLs, domains, "
            f"or local file paths one per line. The list file may have any "
            f"extension; urlCheck verifies it is plain text by inspecting "
            f"its contents.\r\n\r\n"
            f"Output directory: parent directory under which the per-scan "
            f"folders are written. Blank means the current working "
            f"directory.\r\n\r\n"
            f"Options:\r\n"
            f"  Invisible mode - run Edge with no visible browser window\r\n"
            f"  Force replacements - reuse an existing per-page output "
            f"folder (emptying its contents and writing a fresh set of "
            f"files) instead of skipping the URL\r\n"
            f"  View output - open the parent output directory in File "
            f"Explorer when all scans are done\r\n"
            f"  Log session - write urlCheck.log (replacing any prior log) "
            f"in the current working directory\r\n"
            f"  Use configuration - remember these settings for next time "
            f"in %LOCALAPPDATA%\\urlCheck\\urlCheck.ini\r\n\r\n"
            f"Press Cancel to exit without scanning.\r\n\r\n"
            f"Open the full README in your browser?")
        dialogResult = MessageBox.Show(sMsg, f"{sProgramName} - Help",
            MessageBoxButtons.YesNo, MessageBoxIcon.Information,
            MessageBoxDefaultButton.Button2)
        if dialogResult == DialogResult.Yes: launchReadMe()

    def fnOnKeyDown(sender, args):
        if args.KeyCode == Keys.F1:
            args.Handled = True
            args.SuppressKeyPress = True
            fnShowHelp()

    def fnPickFile(sender, args):
        # Pythonnet has a long-standing deadlock when it invokes the modern
        # (Vista+) IFileOpenDialog COM-shell file picker -- documented in
        # pythonnet issues #657 and #1286, both unresolved. The fix is to
        # set AutoUpgradeEnabled = False, which forces the legacy Win32
        # GetOpenFileName common dialog (comdlg32.dll). It has the older
        # Windows look but actually works under pythonnet.
        sCurrent = (txtTarget.Text or "").strip()
        sInitial = getInitialBrowseDir(sCurrent)
        logger.info(f"Browse source clicked; opening OpenFileDialog (legacy) at {sInitial!r}")
        try:
            dialog = OpenFileDialog()
            dialog.AutoUpgradeEnabled = False
            dialog.Title = "Choose a plain text URL list"
            dialog.Filter = ("Plain text files (*.txt;*.lst;*.md)|*.txt;*.lst;*.md|"
                           "All files (*.*)|*.*")
            dialog.FilterIndex = 2  # default to All files; user may have any extension
            dialog.CheckFileExists = True
            dialog.RestoreDirectory = True
            try:
                dialog.InitialDirectory = sInitial
            except Exception: pass
            dialogResult = dialog.ShowDialog()
            logger.info(f"OpenFileDialog returned: {dialogResult}")
            if dialogResult == DialogResult.OK:
                txtTarget.Text = dialog.FileName
                logger.info(f"Source set to: {dialog.FileName}")
        except Exception as ex:
            logger.info(f"OpenFileDialog raised: {ex}")
            MessageBox.Show(f"Browse source failed: {ex}",
                f"{sProgramName} - Browse error",
                MessageBoxButtons.OK, MessageBoxIcon.Warning)

    def fnPickFolder(sender, args):
        # FolderBrowserDialog has no AutoUpgradeEnabled property, and the
        # default shell-COM machinery deadlocks under pythonnet exactly like
        # the modern OpenFileDialog (issue #657). We sidestep it by calling
        # the older Win32 SHBrowseForFolder API directly via ctypes -- which
        # bypasses pythonnet entirely for this dialog and uses the simpler
        # folder picker that doesn't need the same COM-marshaled callbacks.
        sCurrent = (txtOut.Text or "").strip()
        sInitial = getInitialBrowseDir(sCurrent)
        sChosen = ""
        logger.info(f"Choose output clicked; calling SHBrowseForFolder at {sInitial!r}")
        try:
            sChosen = browseForFolderViaShell(
                "Choose the parent directory under which the per-scan "
                "output folder will be created.",
                sInitial)
        except Exception as ex:
            logger.info(f"SHBrowseForFolder raised: {ex}")
            MessageBox.Show(f"Choose output failed: {ex}",
                f"{sProgramName} - Browse error",
                MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        logger.info(f"SHBrowseForFolder returned: {sChosen!r}")
        if sChosen:
            txtOut.Text = sChosen
            logger.info(f"Output dir set to: {sChosen}")

    def fnDefaults(sender, args):
        txtTarget.Text = ""
        txtOut.Text = ""
        chkInvisible.Checked = False
        chkForce.Checked = False
        chkView.Checked = False
        chkLog.Checked = False
        chkUseCfg.Checked = False
        configManager.eraseAll()

    def fnHelpClick(sender, args):
        fnShowHelp()

    def fnOkClick(sender, args):
        sCurrent = (txtTarget.Text or "").strip()
        if not sCurrent:
            MessageBox.Show(
                "Please enter one or more URLs separated by spaces, "
                "or the path to a plain text file of URLs.",
                f"{sProgramName} - Missing source",
                MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtTarget.Focus()
            frm.DialogResult = DialogResult.None_
            return
        sKind, vDetail = classifyInput(sCurrent)
        if sKind == "error":
            MessageBox.Show(str(vDetail),
                f"{sProgramName} - Invalid source",
                MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtTarget.Focus()
            frm.DialogResult = DialogResult.None_
            return
        # sKind is 'urls' or 'listfile' -- both are valid; let the dialog close.

    btnBrowseTarget.Click += EventHandler(fnPickFile)
    btnBrowseOut.Click += EventHandler(fnPickFolder)
    btnDefaults.Click += EventHandler(fnDefaults)
    btnHelp.Click += EventHandler(fnHelpClick)
    btnOk.Click += EventHandler(fnOkClick)
    # KeyDown takes a different EventHandler<KeyEventArgs>; pythonnet handles
    # the conversion when the function signature matches. Pass the function
    # directly rather than wrapping in EventHandler.
    frm.KeyDown += fnOnKeyDown

    txtTarget.Select()
    logger.info("Showing dialog (frm.ShowDialog)")
    dialogResult = frm.ShowDialog()
    logger.info(f"Dialog returned: {dialogResult}")
    if dialogResult != DialogResult.OK:
        frm.Dispose()
        return False

    # Hand values back into the arguments namespace.
    arguments.sSource = (txtTarget.Text or "").strip()
    arguments.sOutputDir = (txtOut.Text or "").strip()
    arguments.bInvisible = bool(chkInvisible.Checked)
    arguments.bForce = bool(chkForce.Checked)
    arguments.bViewOutput = bool(chkView.Checked)
    arguments.bLog = bool(chkLog.Checked)
    arguments.bUseConfig = bool(chkUseCfg.Checked)
    frm.Dispose()
    return True


def launchReadMe():
    """
    Opens README.htm next to urlCheck.exe in the user's default browser.
    Falls back to README.md, then to a polite notice if neither is present.
    """
    sExeDir = ""
    sHtm = ""
    sMd = ""
    sTarget = ""

    sExeDir = os.path.dirname(sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__))
    sHtm = os.path.join(sExeDir, "README.htm")
    sMd = os.path.join(sExeDir, "README.md")
    if os.path.isfile(sHtm): sTarget = sHtm
    elif os.path.isfile(sMd): sTarget = sMd
    if not sTarget:
        try:
            d = _loadDotNetForms()
            if d is not None:
                d["MessageBox"].Show(
                    "Documentation (README.htm or README.md) was not found in:\r\n\r\n" + sExeDir,
                    f"{sProgramName} - Documentation not found",
                    d["MessageBoxButtons"].OK, d["MessageBoxIcon"].Warning)
                return
        except Exception:
            pass
        print(f"[WARN] No README.htm or README.md in {sExeDir}")
        return
    try: os.startfile(sTarget)
    except Exception as ex: print(f"[WARN] Could not open {sTarget}: {ex}")


def showFinalGuiMessage(sText, sTitle):
    """
    Shows a captured-stdout message after a GUI-mode scan completes. Short
    text uses the native MessageBox; long text uses a scrollable read-only
    multi-line TextBox in a small Form. Mirrors 2htm's showFinalMessage.
    """
    bLong = False
    d = None

    d = _loadDotNetForms()
    if d is None:
        print(sText)
        return
    bLong = len(sText) > 800 or sText.count("\n") > 15
    if not bLong:
        d["MessageBox"].Show(sText or "Done. No output.", sTitle,
            d["MessageBoxButtons"].OK, d["MessageBoxIcon"].Information)
        return

    # Long output: scrollable read-only TextBox in a resizable form.
    AnchorStyles = d["AnchorStyles"]
    Button = d["Button"]
    DialogResult = d["DialogResult"]
    DockStyle = d["DockStyle"]
    Font = d["Font"]
    FontFamily = d["FontFamily"]
    Form = d["Form"]
    FormBorderStyle = d["FormBorderStyle"]
    FormStartPosition = d["FormStartPosition"]
    Panel = d["Panel"]
    Point = d["Point"]
    ScrollBars = d["ScrollBars"]
    Size = d["Size"]
    SystemFonts = d["SystemFonts"]
    TextBox = d["TextBox"]

    frm = Form()
    frm.Text = sTitle
    frm.StartPosition = FormStartPosition.CenterScreen
    frm.ClientSize = Size(700, 480)
    frm.FormBorderStyle = FormBorderStyle.Sizable
    frm.MinimizeBox = False
    frm.MaximizeBox = True
    frm.ShowInTaskbar = False
    frm.Font = SystemFonts.MessageBoxFont

    txt = TextBox()
    txt.Multiline = True
    txt.ReadOnly = True
    txt.ScrollBars = ScrollBars.Vertical
    txt.WordWrap = False
    txt.Text = sText
    txt.Dock = DockStyle.Fill
    try: txt.Font = Font(FontFamily.GenericMonospace, 9.0)
    except Exception: pass
    frm.Controls.Add(txt)

    pnl = Panel()
    pnl.Height = 40
    pnl.Dock = DockStyle.Bottom
    frm.Controls.Add(pnl)

    btn = Button()
    btn.Text = "OK"
    btn.DialogResult = DialogResult.OK
    btn.Size = Size(100, 26)
    btn.Anchor = AnchorStyles.Top | AnchorStyles.Right
    btn.Location = Point(pnl.ClientSize.Width - btn.Width - 12, 7)
    pnl.Controls.Add(btn)
    frm.AcceptButton = btn
    frm.CancelButton = btn

    frm.ShowDialog()
    frm.Dispose()


# --- Entry point ---

def main():
    bAutoLaunchedGui = False
    bGuiMode = False
    bMultiUrl = False
    bOwnConsole = False
    iErrorCount = 0
    iSkippedCount = 0
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
    capture = None
    streamOriginalErr = sys.stderr
    streamOriginalOut = sys.stdout
    pathBaseDir = pathlib.Path.cwd()
    pathOutputDir = None
    playwrightCtx = None

    # Playwright suppresses the default SIGINT handler. Restore it so Ctrl+C works.
    signal.signal(signal.SIGINT, signal.SIG_DFL)

    # On Windows, PyInstaller sets the DLL search path to the _MEI temporary folder
    # via SetDllDirectoryW. Child processes (Playwright driver, Edge) inherit this,
    # acquiring handles into _MEI that prevent the bootloader from deleting it on exit.
    # Resetting the DLL directory to NULL before launching Playwright prevents this.
    if sys.platform == "win32":
        ctypes.windll.kernel32.SetDllDirectoryW(None)

    cleanPreviousTempDirs()

    arguments = parseArguments()

    # nargs="*" returns a list. Collapse to a single space-joined string so
    # the rest of main() can treat the field uniformly with the GUI dialog's
    # text-field semantics (one URL, several space-separated URLs, or a path
    # to a URL list file). Empty list -> empty string -> "no input given."
    if isinstance(arguments.sSource, list):
        arguments.sSource = " ".join(arguments.sSource).strip()

    # Open the log file early so the GUI-detection diagnostics can be
    # captured in it. The log is only opened when -l / --log was given on
    # the command line; saved-configuration log preference is honored later
    # in the run, after configuration loading. To debug the auto-detection
    # specifically, run with -l from cmd.exe and again from the desktop
    # shortcut (after editing it to include -l), then compare the two logs.
    if arguments.bLog: logger.open()
    logger.info(f"{sProgramName} {sProgramVersion} starting")
    logger.info(f"Python: {sys.version.split(chr(10))[0]}")
    logger.info(f"Architecture: {struct.calcsize('P') * 8}-bit "
        f"({platform.machine()}); platform={platform.platform()}")
    logger.info(f"Frozen exe: {bool(getattr(sys, 'frozen', False))}; "
        f"executable={sys.executable}")
    if getattr(sys, 'frozen', False):
        # When PyInstaller-bundled, _MEIPASS exposes the temp extraction
        # directory and we can record the bundle's pid for cross-reference.
        logger.info(f"Bundle: _MEIPASS={getattr(sys, '_MEIPASS', '')}; "
            f"pid={os.getpid()}")
    logger.info(f"Working directory: {os.getcwd()}")
    logger.info(f"argv: {sys.argv}")

    # Auto-detect GUI launch via GetConsoleProcessList (primary) with parent-
    # process name as fallback (see isLaunchedFromGui). When invoked with no
    # arguments from a GUI shell -- File Explorer double-click, Start-menu
    # shortcut, desktop hotkey, Run dialog, etc. -- we fall through to GUI
    # mode and hide the otherwise-visible blank console. When invoked from
    # cmd.exe / PowerShell with no arguments we keep CLI behavior and print
    # usage. The -g flag forces GUI mode unconditionally regardless of how
    # the program was launched.
    bOwnConsole = isLaunchedFromGui()
    logger.info(f"bOwnConsole={bOwnConsole}, "
        f"explicit -g={bool(getattr(arguments, 'bGuiMode', False))}, "
        f"sSource={arguments.sSource!r}")
    if not getattr(arguments, "bGuiMode", False):
        if not arguments.sSource and bOwnConsole:
            arguments.bGuiMode = True
            bAutoLaunchedGui = True
            hideOwnConsoleWindow()
    bGuiMode = bool(getattr(arguments, "bGuiMode", False))
    logger.info(f"bGuiMode={bGuiMode}, bAutoLaunchedGui={bAutoLaunchedGui}")

    # Configuration file: opt-in for CLI (-u required), implicit for GUI mode
    # when an existing config file is present. This matches 2htm's asymmetry.
    if bGuiMode and not arguments.bUseConfig and configManager.configExists():
        arguments.bUseConfig = True
    if arguments.bUseConfig:
        configManager.loadInto(arguments)
        # Honor saved log preference even if -l wasn't passed.
        if arguments.bLog and not logger.bEnabled: logger.open()

    # In GUI mode, present the dialog before any other work. The user's
    # choices replace whatever came from CLI and config.
    if bGuiMode:
        if not showGuiDialog(arguments):
            logger.info("User cancelled the dialog")
            logger.close()
            return 0
        # The user may have toggled Log session in the dialog.
        if arguments.bLog and not logger.bEnabled: logger.open()
        # If the user left Use configuration checked, persist their values.
        if arguments.bUseConfig:
            configManager.save(
                arguments.sSource or "",
                arguments.sOutputDir or "",
                arguments.bViewOutput,
                arguments.bInvisible,
                arguments.bForce,
                arguments.bLog)
        logger.info(f"After dialog: input={arguments.sSource!r} "
            f"outputDir={arguments.sOutputDir!r} "
            f"invisible={arguments.bInvisible} "
            f"force={arguments.bForce} "
            f"viewOutput={arguments.bViewOutput} "
            f"log={arguments.bLog} useConfig={arguments.bUseConfig}")

    if not arguments.sSource:
        print(f"{sProgramName} {sProgramVersion}")
        print(sUsage)
        return 1

    # Resolve the base output directory. Default = CWD.
    if arguments.sOutputDir:
        try:
            pathBaseDir = pathlib.Path(arguments.sOutputDir).expanduser()
            pathBaseDir.mkdir(parents=True, exist_ok=True)
        except Exception as ex:
            sErr = f"[ERROR] Output directory '{arguments.sOutputDir}' could not be created: {ex}"
            print(sErr)
            if bGuiMode: showFinalGuiMessage(sErr, f"{sProgramName} - Error")
            return 1
    else:
        pathBaseDir = pathlib.Path.cwd()

    logger.info(f"Output base: {pathBaseDir}")
    logger.info(f"Target: {arguments.sSource}")

    # In GUI mode, capture stdout/stderr so we can present the run summary in
    # a final dialog rather than scrolling past in a console the user can't
    # see. The CLI path leaves stdout/stderr alone.
    if bGuiMode:
        capture = io.StringIO()
        sys.stdout = capture
        sys.stderr = capture

    sInput = str(arguments.sSource).strip()

    # Three-case input dispatch:
    #
    #   (1) sInput is a path to an existing file that is NOT an allowed HTML
    #       extension -> treat as a URL list file (one URL per line).
    #
    #   (2) sInput contains internal whitespace and case (1) did not match
    #       -> treat as a list of URLs separated by spaces. This is what
    #       lets the GUI Source URLs field accept "https://a.com https://b.com".
    #
    #   (3) Otherwise -> single URL / domain / local HTML file path.
    #
    # All three cases produce the same lUrls list, after which the per-URL
    # scan loop below treats every entry uniformly. A multi-URL run (lUrls
    # with more than one element) suppresses the auto-launch of report.htm
    # the same way a URL list file does.
    # Three-case input dispatch using classifyInput():
    #
    #   ('listfile', sPath)  -> read URLs from sPath (.txt only). The list
    #                           file's lines may be URLs, domains, or local
    #                           HTML file paths.
    #   ('urls', lTokens)    -> sInput is one or more space-separated URLs/
    #                           domains. Local HTML file paths are NOT valid
    #                           on direct input and have already been
    #                           rejected by classifyInput.
    #   ('error', sReason)   -> bail with a clear message.
    bMultiUrl = False
    sKind, vDetail = classifyInput(sInput)
    if sKind == "error":
        print(f"[ERROR] {vDetail}")
        logger.info(f"Input error: {vDetail}")
        if bGuiMode:
            sys.stdout = streamOriginalOut
            sys.stderr = streamOriginalErr
            showFinalGuiMessage(capture.getvalue(), f"{sProgramName} - Error")
        logger.close()
        return 1
    if sKind == "listfile":
        try:
            lUrls = getUrlsFromFile(vDetail)
        except Exception as ex:
            print(f"[ERROR] Could not read URL list file: {ex}")
            if bGuiMode:
                sys.stdout = streamOriginalOut
                sys.stderr = streamOriginalErr
                showFinalGuiMessage(capture.getvalue(), f"{sProgramName} - Error")
            logger.close()
            return 1
        bMultiUrl = True
        iUrlTotal = len(lUrls)
        logger.info(f"URL list: {vDetail} ({iUrlTotal} URL(s))")
    else:
        # sKind == "urls"
        lUrls = [getNormalizedUrl(sToken) for sToken in vDetail if sToken]
        iUrlTotal = len(lUrls)
        bMultiUrl = iUrlTotal > 1
        if bMultiUrl: logger.info(f"URL set: {iUrlTotal} URL(s) from the source field")

    try:
        with sync_playwright() as playwrightCtx:
            browserType = playwrightCtx.chromium
            lArgs = [
                "--mute-audio",
                "--no-default-browser-check",
                "--no-first-run",
                f"--window-size={iDefaultViewportWidth},{iDefaultViewportHeight}",
            ]
            # Playwright still calls this "headless"; we expose it to the
            # user as "Invisible" because that wording is clearer and not
            # tied to the implementation. The mapping is a 1:1 boolean.
            #
            # urlCheck drives the system-installed Microsoft Edge through
            # channel="msedge"; we never download a Playwright-bundled
            # Chromium. On modern Windows 10/11, Edge ships in-box, so
            # this almost always succeeds. If the user has somehow removed
            # Edge or is on an unusual configuration where Playwright
            # cannot find it, surface a friendly message rather than a
            # raw Python traceback.
            try:
                browser = browserType.launch(channel=sBrowserChannel, headless=bool(arguments.bInvisible), args=lArgs)
            except Exception as ex:
                sErr = (
                    f"Could not launch Microsoft Edge: {ex}\n\n"
                    "urlCheck requires Microsoft Edge, which ships with "
                    "Windows 10 and 11 by default. If Edge has been "
                    "removed or is unavailable, install or repair it "
                    "from https://www.microsoft.com/edge and try again."
                )
                print(f"[ERROR] {sErr}")
                logger.info(f"Browser launch failed: {ex}")
                if bGuiMode:
                    sys.stdout = streamOriginalOut
                    sys.stderr = streamOriginalErr
                    showFinalGuiMessage(capture.getvalue(), f"{sProgramName} - Edge not available")
                logger.close()
                return 1
            context = browser.new_context(bypass_csp=True, ignore_https_errors=bDefaultIgnoreHttpsErrors, user_agent=sUserAgent, viewport={"width": iDefaultViewportWidth, "height": iDefaultViewportHeight})
            # Pre-fetch axe-core content once so CSP-restricted sites can still be scanned.
            sAxeContent = ""
            for sAxeUrl in aAxeCdnUrls:
                try:
                    sAxeContent = fetchText(sAxeUrl)
                    logger.info(f"axe-core pre-fetched from {sAxeUrl}")
                    break
                except Exception:
                    continue
            iUrlIndex = 0
            for sUrl in lUrls:
                iUrlIndex += 1
                # lUrls is already a list of normalized targets regardless of
                # which dispatch case produced it; we pass each entry through
                # scanUrl unchanged. (We re-run getNormalizedUrl idempotently
                # for the listfile case where lines were read raw from the
                # file without prior normalization.)
                sNormalizedUrl = getNormalizedUrl(sUrl) if sKind == "listfile" else sUrl
                if iUrlTotal > 1: logger.info(f"[{iUrlIndex}/{iUrlTotal}] {sNormalizedUrl}")
                try:
                    vResult = scanUrl(
                        sUrl, sNormalizedUrl, browser, context, pathBaseDir,
                        sAxeContent, bForce=bool(arguments.bForce))
                    if vResult == "skipped":
                        iSkippedCount += 1
                except Exception as ex:
                    iErrorCount += 1
                    # Surface the error to the console (always, since
                    # console output is the user's expected view of run
                    # results) and to the log (only if -l is on; the
                    # logger silently no-ops otherwise). We never create
                    # an error file on disk that the user did not
                    # explicitly request.
                    print(f"Error scanning {sNormalizedUrl}: {ex}")
                    logger.info(f"Error scanning {sNormalizedUrl}: {ex}")
                    logger.info(f"Traceback:\n{traceback.format_exc()}")
            if iUrlTotal > 1:
                iSucceeded = iUrlTotal - iErrorCount - iSkippedCount
                sSummary = f"Done. {iSucceeded} of {iUrlTotal} URLs scanned successfully"
                if iSkippedCount > 0: sSummary += f", {iSkippedCount} skipped"
                if iErrorCount > 0: sSummary += f", {iErrorCount} error(s)"
                sSummary += "."
                logger.info(sSummary)
                print(sSummary)
                if iErrorCount > 0 and arguments.bLog:
                    logger.info("Error details (per-URL traceback) above in this log.")
                elif iErrorCount > 0:
                    print("Re-run with -l to log full error tracebacks to urlCheck.log.")
            # Open the parent output directory once at the end of the run, if
            # requested. This shows the user all per-page subdirectories at
            # once rather than focusing on any single page's folder.
            if arguments.bViewOutput: openFolderInExplorer(str(pathBaseDir))
    finally:
        try:
            if context is not None: context.close()
        except Exception:
            pass
        try:
            if browser is not None: browser.close()
        except Exception:
            pass

        # Restore stdout/stderr and surface captured output in a GUI dialog.
        if bGuiMode and capture is not None:
            sys.stdout = streamOriginalOut
            sys.stderr = streamOriginalErr
            sCaptured = capture.getvalue()
            sTitle = f"{sProgramName} - Results" if iErrorCount == 0 else f"{sProgramName} - Completed with errors"
            showFinalGuiMessage(sCaptured if sCaptured else "Done. No output.", sTitle)

        logger.info(f"Done. {iUrlTotal - iErrorCount} of {iUrlTotal} scanned. {iErrorCount} error(s).")
        logger.close()

    return 0 if iErrorCount == 0 else 1


if __name__ == "__main__": raise SystemExit(main())

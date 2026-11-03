"""
Bug_Report_Sumaiya_Nasrin.py
SQA assessment - Supporters Frame (https://supportersframe.com/)
Prepared by: Sumaiya Nasrin

SECTION A  Bug data (single source of truth for the app AND the Excel file)
SECTION B  Excel export (openpyxl)
SECTION C  Desktop application (Tkinter)

Run the app:            python Bug_Report_Sumaiya_Nasrin.py
Create Excel only:      python Bug_Report_Sumaiya_Nasrin.py --export
Screenshots are read from the folder named "evidence" next to this file
(or next to the .exe). Name them BUG-001_something.png, BUG-002_something.png ...
"""

import math
import os
import sys
from collections import Counter

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

try:  # Pillow is only needed to embed screenshots into the Excel file
    import PIL  # noqa: F401
    from openpyxl.drawing.image import Image as XLImage
    PIL_OK = True
except Exception:  # pragma: no cover
    PIL_OK = False

try:  # Tkinter is only needed for the window, not for --export
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    TK_OK = True
except Exception:  # pragma: no cover
    TK_OK = False


# ======================================================================
# SECTION A - DATA
# ======================================================================

AUTHOR = "Sumaiya Nasrin"
WEBSITE = "https://supportersframe.com/"
CREATE_URL = "https://supportersframe.com/create/name-twibbon"
TEST_DATE = "2-3 November 2026"
ENVIRONMENT = "Windows PC, Google Chrome 156.0.8078.76 (64-bit), Chrome DevTools, Lighthouse"

LIGHTHOUSE = {
    "Audited page": "Homepage - https://supportersframe.com/",
    "Device": "Mobile (simulated), Navigation mode",
    "Performance": 26,
    "Accessibility": 84,
    "Best Practices": 73,
    "SEO": 92,
}

BUGS = [
    {
        "id": "BUG-001",
        "title": "Navigation stylesheet nav.css returns HTTP 404 on the campaign creation page",
        "module": "Page resources / Campaign creation",
        "preconditions": "None. Guest user, Chrome with DevTools.",
        "steps": [
            f"Open {CREATE_URL} in Chrome.",
            "Press F12, open the Network tab and tick 'Disable cache'.",
            "Reload the page with F5.",
            "Type 'nav.css' in the Network filter box and read the Status column.",
            "Open the Console tab and filter by 'nav.css'.",
        ],
        "expected": "The stylesheet loads successfully (HTTP 200) and no error appears in the Console.",
        "actual": ("GET https://supportersframe.com/create/nav.css returns 404 (Not Found). "
                   "Initiator: name-twibbon:40. The Response tab shows 'No data found for "
                   "resource with given identifier' and the Console logs the failed request. "
                   "The page appeared normally styled, so no visible breakage was confirmed."),
        "severity": "Low",
        "severity_reason": "A resource fails to load, but no visible or functional impact was demonstrated.",
        "evidence": ["BUG-001_network_404.png", "BUG-001_console_404.png"],
        "reproducibility": "Reproduced on repeated reloads.",
        "status": "Open",
    },
    {
        "id": "BUG-002",
        "title": "JavaScript SyntaxError on the campaign creation page (name-twibbon, line 4778)",
        "module": "Campaign creation / Frontend script",
        "preconditions": "None. Guest user, Chrome with DevTools.",
        "steps": [
            f"Open {CREATE_URL} in Chrome.",
            "Press F12 and open the Console tab.",
            "Reload the page with F5.",
            "Show only the 'Errors' level and read the red message.",
        ],
        "expected": "The page scripts load without JavaScript syntax errors.",
        "actual": ("The Console shows: Uncaught SyntaxError: Unexpected token ')' "
                   "(at name-twibbon:4778:13). The error appears after each reload. "
                   "Frame upload, text tool and preview still worked in testing, so the "
                   "user-visible impact is not confirmed."),
        "severity": "Medium",
        "severity_reason": ("A script block on the page fails to parse, which may disable some "
                            "functionality. The exact impact was not confirmed."),
        "evidence": ["BUG-002_console_syntaxerror.png"],
        "reproducibility": "Reproduced after reload.",
        "status": "Open",
    },
    {
        "id": "BUG-003",
        "title": "Empty or whitespace-only campaign title is accepted without any validation message",
        "module": "Campaign creation / Campaign Details",
        "preconditions": "A valid transparent PNG frame is uploaded on step 1. User is not logged in.",
        "steps": [
            f"Open {CREATE_URL}.",
            "Upload a valid PNG frame with a transparent center.",
            "Leave 'Campaign Title' empty (test 1), or enter only spaces (test 2).",
            "Click Next and continue to the publish step.",
        ],
        "expected": ("The title is required: a validation message appears and the user cannot "
                     "continue until a real title is entered."),
        "actual": ("No message appeared and Next moved forward in both tests. With an empty title "
                   "the campaign was displayed as 'Untitled campaign'. The flow stopped at a login "
                   "prompt, so whether a campaign would be saved was not verified, and the display "
                   "result for the spaces-only title was not verified."),
        "severity": "Low",
        "severity_reason": "The site falls back to a default name and nothing breaks.",
        "evidence": ["BUG-003_empty_title_flow.png"],
        "reproducibility": "Observed in 1 test run per input; repeat recommended.",
        "status": "Open",
    },
    {
        "id": "BUG-004",
        "title": "Very long campaign title overflows the publish dialog; no maximum title length is enforced",
        "module": "Campaign creation / Campaign Details and publish dialog",
        "preconditions": "A valid transparent PNG frame is uploaded. User is not logged in.",
        "steps": [
            f"Open {CREATE_URL} and upload a valid frame.",
            "In 'Campaign Title', enter a string of more than 300 characters (for example the letter 'a' repeated).",
            "Click Next and continue to the publish step.",
            "Look at how the title is displayed in the publish dialog.",
        ],
        "expected": ("The title length is limited (with a counter or a message), or the long title "
                     "is wrapped or truncated so it stays inside the dialog."),
        "actual": ("The 300+ character title was accepted with no message. In the publish dialog "
                   "the title is shown on a single line and runs past the right edge of the dialog "
                   "until it is cut off at the edge of the screen. The Description field shows a "
                   "character counter (e.g. 4/250), but no counter was visible on the title field."),
        "severity": "Low",
        "severity_reason": ("Cosmetic layout overflow; the Publish button stays usable and the user "
                            "can shorten the title."),
        "evidence": ["BUG-004_long_title.png"],
        "reproducibility": "Observed in 1 test run; repeat recommended.",
        "status": "Open",
    },
    {
        "id": "BUG-005",
        "title": "Cancel button discards all entered campaign data without confirmation",
        "module": "Campaign creation / Navigation",
        "preconditions": "Campaign details are entered and a frame is uploaded.",
        "steps": [
            f"Open {CREATE_URL}.",
            "Upload a frame and enter a campaign title and description.",
            "Click 'Cancel' on the first page.",
        ],
        "expected": ("A confirmation dialog (for example 'Are you sure? All entered information "
                     "will be lost') lets the user continue editing or confirm."),
        "actual": "All entered information was cleared immediately and no confirmation dialog appeared.",
        "severity": "Low",
        "severity_reason": "Data loss follows a deliberate click and the amount of data is small.",
        "evidence": ["BUG-005_cancel_before.png", "BUG-005_cancel_after.png"],
        "reproducibility": "Observed in 1 test run; repeat recommended.",
        "status": "Open",
    },
    {
        "id": "BUG-006",
        "title": "Frame upload error message gives no actionable instructions for fixing the image",
        "module": "Campaign creation / Frame upload",
        "preconditions": "User is on the frame upload step.",
        "steps": [
            f"Open {CREATE_URL}.",
            "Upload a PNG or JPG image that has no transparent area.",
            "Read the pop-up that appears.",
        ],
        "expected": ("A plain-language message explains the problem and how to fix it, for example: "
                     "'Your image has no transparent area. Please upload a PNG frame with a "
                     "transparent center.'"),
        "actual": ("The pop-up says 'No transparent area detected. This means that there is no hole "
                   "on your frame. To give the necessary hole, you can follow this Help Center'. "
                   "It uses informal wording ('hole') and does not say what file to upload; it "
                   "only links to the general Help Center."),
        "severity": "Low",
        "severity_reason": ("The invalid file is correctly rejected and a help link exists, but the "
                            "user is not told how to fix it."),
        "evidence": ["BUG-006_error_message.png"],
        "reproducibility": "Reproduced with a no-transparency PNG and a JPG.",
        "status": "Open",
    },
    {
        "id": "BUG-007",
        "title": "Empty (0 KB) file named .png is rejected with the misleading 'No transparent area detected'",
        "module": "Campaign creation / Frame upload",
        "preconditions": "A 0 KB file named faked.png exists on the computer.",
        "steps": [
            f"Open {CREATE_URL}.",
            "Click the frame upload area.",
            "Select faked.png (size 0 KB) and click Open.",
            "Read the pop-up.",
        ],
        "expected": "An error that names the real problem, e.g. 'This file is not a valid image.'",
        "actual": ("The same pop-up as for a valid-but-opaque image appears: 'No transparent area "
                   "detected. This means that there is no hole on your frame...'. The message "
                   "names a cause that does not match the real problem."),
        "severity": "Low",
        "severity_reason": "Rejection is correct, but the reported cause is misleading.",
        "evidence": ["BUG-007_picker.png", "BUG-007_message.png"],
        "reproducibility": "Observed in 1 test run; repeat recommended.",
        "status": "Open",
    },
    {
        "id": "BUG-008",
        "title": "Mobile search and hamburger menu buttons have no accessible name",
        "module": "Homepage / Navigation (accessibility)",
        "preconditions": "Chrome Lighthouse, Mobile mode, Navigation, Incognito window.",
        "steps": [
            f"Open {WEBSITE} in an Incognito window.",
            "Press F12, open the Lighthouse tab, choose Mobile and Accessibility.",
            "Click 'Analyze page load'.",
            "Open the failed audit 'Buttons do not have an accessible name'.",
        ],
        "expected": ("Icon-only buttons have accessible names such as 'Search' and "
                     "'Open navigation menu'."),
        "actual": ("Lighthouse reports button#mobileSearchBtn and button#mobile-hamburger as "
                   "having no accessible name; screen readers would announce only 'button'. "
                   "In the Elements tab the search button contains only an icon and has no "
                   "aria-label or title."),
        "severity": "Medium",
        "severity_reason": "Screen-reader users cannot identify the purpose of core navigation controls.",
        "evidence": ["BUG-008_lighthouse_buttons.png"],
        "reproducibility": "Reproduced in the Lighthouse run.",
        "status": "Open",
    },
    {
        "id": "BUG-009",
        "title": "Insufficient color contrast on gray text and footer links on the homepage",
        "module": "Homepage / UI (accessibility)",
        "preconditions": "Chrome Lighthouse, Mobile mode, Incognito window.",
        "steps": [
            f"Open {WEBSITE} in an Incognito window.",
            "Press F12, open Lighthouse, choose Mobile and Accessibility, and click 'Analyze page load'.",
            "Open the failed audit 'Background and foreground colors do not have a sufficient contrast ratio'.",
        ],
        "expected": "Text meets WCAG contrast requirements against its background.",
        "actual": ("Lighthouse flags gray text (text-gray-400 paragraphs), number labels #1 to #9, "
                   "the 'Create Your Free Frame' button, the footer description, and the Privacy "
                   "Policy, Refund Policy, Terms & Condition links and copyright text."),
        "severity": "Medium",
        "severity_reason": "Low-contrast text is hard to read for users with low vision.",
        "evidence": ["BUG-009_lighthouse_contrast.png"],
        "reproducibility": "Reproduced in the Lighthouse run.",
        "status": "Open",
    },
    {
        "id": "BUG-010",
        "title": "Heading elements are not in a sequentially-descending order on the homepage",
        "module": "Homepage / Page structure (accessibility)",
        "preconditions": "Chrome Lighthouse, Mobile mode, Incognito window.",
        "steps": [
            f"Open {WEBSITE} in an Incognito window.",
            "Press F12, open Lighthouse, choose Mobile and Accessibility, and click 'Analyze page load'.",
            "Open the failed audit 'Heading elements are not in a sequentially-descending order'.",
        ],
        "expected": "Heading levels follow a logical order without skipping levels (H1, then H2, then H3).",
        "actual": ("Lighthouse flags heading levels that skip the expected order, including an h4 "
                   "(class 'text-2xl font-light'), an h5 (class 'banner-creators__animation-container') "
                   "and an h3 (class 'card-title')."),
        "severity": "Low",
        "severity_reason": ("It makes the page outline harder to follow for screen-reader users, "
                            "but the content stays readable."),
        "evidence": ["BUG-010_lighthouse_heading_order.png"],
        "reproducibility": "Reproduced in the Lighthouse run.",
        "status": "Open",
    },
    {
        "id": "BUG-011",
        "title": "Heavy page weight: about 8.8 MB transferred over about 125 requests",
        "module": "Performance",
        "preconditions": "Incognito window, DevTools Network tab, 'Disable cache' ticked.",
        "steps": [
            f"Open {WEBSITE} in an Incognito window.",
            "Press F12, open the Network tab and tick 'Disable cache'.",
            "Reload with F5 and read the summary bar at the bottom.",
            "Sort by Size to see the largest files.",
        ],
        "expected": "A reasonably light page with compressed, appropriately sized images.",
        "actual": ("Two runs showed 125 and 126 requests and 8.8 MB transferred, with Load times of "
                   "4.14 s and 3.79 s. Several PNG images were about 1.4 MB each. For comparison, "
                   "en.wikipedia.org showed 32 requests and 766 KB (different content, so this is a "
                   "reference only). Mobile Lighthouse Performance score: 26 (simulated slow device)."),
        "severity": "Low",
        "severity_reason": ("The site works, but the large payload wastes data and slows loading "
                            "on slow connections."),
        "evidence": ["BUG-011_network_summary.png"],
        "reproducibility": "Reproduced in 2 of 2 later runs (page weight).",
        "status": "Open",
    },
]

# Findings that are NOT counted as bugs (kept for transparency)
OBSERVATIONS = [
    {
        "id": "OBS-001",
        "title": "Document does not have a main landmark (homepage)",
        "status": "Not reproduced / Previously observed",
        "note": ("Flagged by Lighthouse in an earlier homepage audit; not reproduced in the final "
                 "verification, so it is not counted as a confirmed bug."),
    },
    {
        "id": "OBS-002",
        "title": "One very slow page load (41.88 s, about 17 s document response)",
        "status": "Not reproduced",
        "note": ("Recorded in a single first run. Two later runs loaded in about 4 s, so this is "
                 "treated as an intermittent event and not reported as a defect."),
    },
    {
        "id": "OBS-003",
        "title": "Console warnings: Swiper loop warning, Tailwind CDN warning, Google Ads connection error",
        "status": "Not a bug",
        "note": "Third-party or best-practice warnings with no visible effect on the tested pages.",
    },
]

COVERAGE_NOTE = (
    "Unable to repeat website tests on 3 November 2026: the homepage returned an inactive-store/"
    "forbidden (403) page and the campaign route returned 404 in the available browser. Findings "
    "below are based on the dated original tests and screenshots; they have not been independently "
    "reverified today."
)

COVERAGE = [
    (
        "Homepage",
        "Partial",
        "Homepage mobile Lighthouse accessibility audit and network/performance observations.",
        "Current retest returned an inactive-store/forbidden (403) page; not a complete functional review of links, forms, or content.",
    ),
    (
        "Campaign creation",
        "Partial",
        "Frame upload, title/link/description inputs, back/cancel navigation, and browser console/network.",
        "Current retest returned 404. Earlier flow reached a login prompt; authenticated publishing was not verified.",
    ),
    (
        "Accessibility",
        "Partial",
        "Mobile Lighthouse findings for button names, color contrast, and heading order.",
        "Manual keyboard-only and screen-reader testing were not performed.",
    ),
    (
        "Performance",
        "Partial",
        "Two later homepage network runs and a mobile Lighthouse run.",
        "No full-page inventory, desktop performance run, or real-device measurement.",
    ),
    (
        "Responsive layouts",
        "Limited",
        "Homepage checked with Lighthouse mobile emulation.",
        "Dedicated viewport-by-viewport layout checks and tablet/real-device tests were not performed.",
    ),
    (
        "Browser compatibility",
        "Not tested",
        "Chrome was the recorded test browser.",
        "Microsoft Edge and Firefox were not tested.",
    ),
    (
        "Authentication and publishing",
        "Not tested",
        "The unauthenticated flow was followed to the login prompt.",
        "Authenticated account, save, and final publish behavior were not verified.",
    ),
    (
        "Other site pages and routes",
        "Not tested",
        "The recorded scope focuses on the homepage and campaign creation page.",
        "Other pages, navigation destinations, and site-wide workflows were not systematically explored.",
    ),
    (
        "Retest of previous findings",
        "Blocked",
        "Attempted homepage and campaign-route verification on 3 November 2026.",
        "Homepage returned inactive-store/forbidden (403); campaign route returned 404. Repeatability of single-run findings remains unconfirmed.",
    ),
]

RETEST_STATUS = (
    "Not run - site currently returns 404/403; execute when the site is available"
)

RETEST_PLAN = [
    (
        "RT-01", "P1", "Availability and route smoke test",
        "Chrome and Edge; desktop",
        "Open the homepage and campaign creation URL; follow the primary navigation links.",
        "Pages load without an inactive-store, 403, or unexpected 404 response; links reach the intended pages.",
    ),
    (
        "RT-02", "P1", "Repeat existing defects",
        "Chrome; desktop and mobile emulation",
        "Repeat BUG-001 through BUG-010 steps three times each where applicable; record each run separately.",
        "Record whether each issue reproduces in 0/3, 1/3, 2/3, or 3/3 runs; update severity only from observed impact.",
    ),
    (
        "RT-03", "P1", "End-to-end campaign publishing",
        "Chrome; desktop and mobile",
        "With an authorized test account, create a campaign, complete required fields, and publish it.",
        "Campaign is saved once, confirmation is clear, and the published campaign can be opened; do not use a personal account without authorization.",
    ),
    (
        "RT-04", "P1", "Responsive viewport sweep",
        "Chrome responsive emulation: 320, 375, 390, 768, 1024, and 1440 CSS px",
        "Check homepage and campaign creation at each width; inspect horizontal overflow, clipped controls, dialogs, and navigation.",
        "No unintended horizontal scrolling or clipped content; controls and dialogs remain visible and usable.",
    ),
    (
        "RT-05", "P1", "Cross-browser smoke test",
        "Chrome and Edge; Firefox if available",
        "Repeat homepage navigation, campaign form entry, frame upload, back/cancel, and preview.",
        "Core workflows and layout behave consistently; log browser/version and any browser-specific differences.",
    ),
    (
        "RT-06", "P1", "Campaign title boundaries",
        "Chrome; desktop and mobile",
        "Test empty, spaces-only, one-character, 250-character, and 301-character titles; inspect validation and preview.",
        "Required-value rules are clear; long text does not obscure controls or escape the dialog.",
    ),
    (
        "RT-07", "P1", "Frame upload validation",
        "Chrome; desktop",
        "Test valid transparent PNG, opaque PNG, JPG, zero-byte file, renamed non-image, and a file above the stated size limit if one exists.",
        "Each file is accepted or rejected for the correct reason; messages identify the actual problem and next step.",
    ),
    (
        "RT-08", "P1", "Keyboard-only operation",
        "Chrome; desktop",
        "Use Tab, Shift+Tab, Enter, Space, and Escape through navigation, upload, form, dialogs, and publish steps.",
        "Focus is visible and ordered logically; every control is reachable and operable without a pointer.",
    ),
    (
        "RT-09", "P1", "Accessibility audit and manual semantics",
        "Chrome Lighthouse mobile and desktop; keyboard review",
        "Rerun accessibility audits; inspect accessible names, contrast, heading hierarchy, landmarks, and focus order.",
        "Record audit scores and affected elements; verify manual findings independently instead of relying only on a score.",
    ),
    (
        "RT-10", "P2", "Cancel and state retention",
        "Chrome; desktop and mobile",
        "Enter frame and campaign details, use Back, then Cancel; repeat with browser refresh where safe.",
        "Back retains expected input; Cancel clearly warns before destructive loss or the product requirement is documented.",
    ),
    (
        "RT-11", "P2", "Network and performance repeat",
        "Chrome; desktop and mobile emulation",
        "Run three uncached homepage loads and three campaign-page loads; capture request count, transferred bytes, and load timing.",
        "Report medians and outliers per route/device; distinguish first-party failures from third-party requests.",
    ),
    (
        "RT-12", "P2", "Site-wide page and link inventory",
        "Chrome; desktop",
        "Visit each page linked from primary navigation and footer; exercise visible forms and key calls to action.",
        "Every tested route loads and each link reaches its labeled destination; record unvisited routes explicitly.",
    ),
]

TEST_LOG = [
    ("T1", "Campaign title", "Empty title", "Accepted; shown as 'Untitled campaign'", "BUG-003"),
    ("T2", "Campaign title", "Spaces-only title", "Accepted in one run; display not verified; retest pending", "BUG-003"),
    ("T3", "Campaign title", "Very short title", "Accepted - Passed", "-"),
    ("T4", "Campaign title", "Title over 300 characters", "Overflow observed once; retest pending", "BUG-004"),
    ("T5", "Campaign title", "HTML-like title: QA Test <b>bold</b>", "Displayed as plain text - Passed", "-"),
    ("T6", "Description", "Empty description", "Accepted; not treated as a defect", "-"),
    ("T7", "Campaign link", "Link 'abc'", "Message 'This link is already taken' - Passed", "-"),
    ("T8", "Campaign link", "Link 'www.example.com'", "Validation message shown - Passed", "-"),
    ("T9", "Campaign link", "Link 'http://'", "Validation message shown - Passed", "-"),
    ("T10", "Campaign link", "Empty link", "Validation message shown - Passed", "-"),
    ("U1", "Frame upload", "PNG with transparent center", "Accepted - Passed", "-"),
    ("U2", "Frame upload", "PNG without transparency", "Rejected with 'No transparent area detected'", "BUG-006"),
    ("U3", "Frame upload", "JPG image", "Rejected with the same message", "BUG-006"),
    ("U4", "Frame upload", "Empty 0 KB file named faked.png", "Misleading rejection observed once; retest pending", "BUG-007"),
    ("U5", "Frame upload", "Very large image (over 10 MB)", "Not tested (the file used was 1.07 MB)", "-"),
    ("N1", "Navigation", "Back button keeps entered data", "Passed", "-"),
    ("N2", "Navigation", "Cancel button", "Cleared data without confirmation in one run; retest pending", "BUG-005"),
    ("N3", "Navigation", "Repeated clicks on the final submit button", "Not tested", "-"),
    ("P1", "Performance", "Network runs and Wikipedia control", "Heavy page weight", "BUG-011"),
    ("A1", "Accessibility", "Lighthouse mobile audit of the homepage", "Failed audits found", "BUG-008, 009, 010"),
    ("C1", "Compatibility", "Microsoft Edge and Firefox", "Not tested", "-"),
    ("R1", "Responsive", "Device toolbar (mobile and tablet)", "Not included in this report", "-"),
    ("R2", "Retest", "Repeat previously reported website findings", "Blocked on 3 November 2026: site returned inactive-store/forbidden (403)", "-"),
]

SEVERITIES = ["High", "Medium", "Low"]
STATUSES = ["Open", "Fixed", "Closed", "Not a bug"]
IMAGE_EXTENSIONS = (".png", ".jpg", ".jpeg", ".gif", ".bmp")


def app_dir():
    """Folder of the .py file, or of the .exe when frozen by PyInstaller."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def evidence_dir():
    bundle_dir = getattr(sys, "_MEIPASS", app_dir())
    return os.path.join(bundle_dir, "evidence")


def steps_as_text(bug):
    return "\n".join(f"{i}. {s}" for i, s in enumerate(bug["steps"], start=1))


def evidence_exists(name):
    return os.path.isfile(os.path.join(evidence_dir(), name))


def missing_evidence():
    """Evidence file names that do not exist in the evidence folder."""
    return [n for b in BUGS for n in b["evidence"] if not evidence_exists(n)]


# ======================================================================
# SECTION B - EXCEL EXPORT
# ======================================================================

FONT_NAME = "Arial"
HEADERS = [
    ("Bug ID", 10), ("Bug Title / Summary", 34), ("Module / Feature", 22),
    ("Preconditions", 26), ("Steps to Reproduce", 46), ("Expected Result", 34),
    ("Actual Result", 52), ("Severity", 11), ("Severity Reason", 32),
    ("Test Environment", 24), ("Screenshot / Evidence", 32),
    ("Reproducibility", 24), ("Status", 10),
]
SEVERITY_FILL = {
    "High": PatternFill("solid", fgColor="F8CBAD"),
    "Medium": PatternFill("solid", fgColor="FFE699"),
    "Low": PatternFill("solid", fgColor="C6E0B4"),
}
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
BAND_FILL = PatternFill("solid", fgColor="DDEBF7")
THIN = Side(style="thin", color="999999")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _font(**kw):
    kw.setdefault("size", 10)
    return Font(name=FONT_NAME, **kw)


def _estimate_height(pairs):
    """openpyxl cannot auto-fit rows, so estimate the row height from the text."""
    max_lines = 1
    for text, width in pairs:
        per_line = max(int(width * 1.05), 1)
        lines = sum(max(1, math.ceil(len(part) / per_line)) for part in str(text).split("\n"))
        max_lines = max(max_lines, lines)
    return min(max_lines * 13.5 + 8, 400)


def _header_row(ws, row, labels):
    for col, label in enumerate(labels, start=1):
        c = ws.cell(row=row, column=col, value=label)
        c.font = _font(bold=True, color="FFFFFF")
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER


def _fill_evidence_sheet(ws, bugs):
    """Embed the screenshots. Returns {file name: 'embedded' | 'missing' | ...}."""
    results = {}
    ws.column_dimensions["A"].width = 125
    ws["A1"] = "Evidence Screenshots"
    ws["A1"].font = _font(bold=True, size=14, color="1F4E78")
    row = 3
    for bug in bugs:
        h = ws.cell(row=row, column=1, value=f"{bug['id']} - {bug['title']}")
        h.font = _font(bold=True, color="FFFFFF")
        h.fill = HEADER_FILL
        row += 1
        for name in bug["evidence"]:
            ws.cell(row=row, column=1, value=name).font = _font(italic=True, size=9)
            row += 1
            path = os.path.join(evidence_dir(), name)
            if not os.path.isfile(path):
                ws.cell(row=row, column=1, value="[Screenshot file not found when this report was exported]").font = _font(color="C00000")
                results[name] = "missing"
                row += 2
                continue
            if not PIL_OK:
                ws.cell(row=row, column=1, value="[Pillow is not installed, so the image was not embedded]").font = _font(color="C00000")
                results[name] = "no_pillow"
                row += 2
                continue
            try:
                img = XLImage(path)
                w, h_px = img.width, img.height
                scale = min(1.0, 820 / w, 520 / h_px)
                img.width, img.height = int(w * scale), int(h_px * scale)
                ws.add_image(img, f"A{row}")
                row += math.ceil(img.height / 20) + 2
                results[name] = "embedded"
            except Exception as exc:  # corrupted or unreadable image
                ws.cell(row=row, column=1, value=f"[Could not embed this image: {exc}]").font = _font(color="C00000")
                results[name] = "error"
                row += 2
        row += 1
    return results


def _evidence_text(bug, results):
    tags = {"embedded": "(see 'Evidence' sheet)", "missing": "(file not found)",
            "no_pillow": "(not embedded)", "error": "(could not embed)"}
    return "\n".join(f"{n} {tags.get(results.get(n), '')}".strip() for n in bug["evidence"])


def _count_table(ws, top, title, items, counts, colour_by_severity=False):
    ws.cell(row=top, column=1, value=title).font = _font(bold=True, size=12, color="1F4E78")
    _header_row(ws, top + 1, ["Category", "Count"])
    r = top + 2
    for item in items:
        a = ws.cell(row=r, column=1, value=item)
        a.font, a.border = _font(), BORDER
        if colour_by_severity:
            a.fill = SEVERITY_FILL[item]
        n = ws.cell(row=r, column=2, value=counts.get(item, 0))
        n.font, n.border = _font(), BORDER
        n.alignment = Alignment(horizontal="center")
        r += 1
    t = ws.cell(row=r, column=1, value="Total")
    t.font, t.border = _font(bold=True), BORDER
    tc = ws.cell(row=r, column=2, value=sum(counts.get(i, 0) for i in items))
    tc.font, tc.border = _font(bold=True), BORDER
    tc.alignment = Alignment(horizontal="center")
    return r + 2


def export_to_excel(path, bugs=None):
    """Create the workbook. Returns a dict describing the screenshots."""
    bugs = BUGS if bugs is None else bugs
    wb = Workbook()
    wb.properties.creator = AUTHOR
    wb.properties.title = "SQA Bug Report - Supporters Frame"

    ws = wb.active
    ws.title = "Bug Report"
    sm = wb.create_sheet("Summary")
    tl = wb.create_sheet("Test Log")
    ev = wb.create_sheet("Evidence")
    coverage = wb.create_sheet("Coverage")
    retest = wb.create_sheet("Retest Plan")

    results = _fill_evidence_sheet(ev, bugs)

    # ---- Sheet 1: Bug Report ----
    _header_row(ws, 1, [h for h, _ in HEADERS])
    for col, (_, width) in enumerate(HEADERS, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.row_dimensions[1].height = 32
    for row, bug in enumerate(bugs, start=2):
        values = [
            bug["id"], bug["title"], bug["module"], bug["preconditions"],
            steps_as_text(bug), bug["expected"], bug["actual"], bug["severity"],
            bug["severity_reason"], ENVIRONMENT, _evidence_text(bug, results),
            bug["reproducibility"], bug["status"],
        ]
        for col, value in enumerate(values, start=1):
            c = ws.cell(row=row, column=col, value=value)
            c.font = _font()
            c.border = BORDER
            c.alignment = Alignment(vertical="top", wrap_text=True,
                                    horizontal="center" if col in (1, 8, 13) else "left")
        sev = ws.cell(row=row, column=8)
        sev.fill = SEVERITY_FILL.get(bug["severity"], PatternFill())
        sev.font = _font(bold=True)
        ws.row_dimensions[row].height = _estimate_height(
            [(v, HEADERS[i][1]) for i, v in enumerate(values)])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}{len(bugs) + 1}"
    ws.sheet_view.zoomScale = 90
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # ---- Sheet 2: Summary ----
    sm.column_dimensions["A"].width = 30
    sm.column_dimensions["B"].width = 14
    for letter in "CDEFG":
        sm.column_dimensions[letter].width = 16
    sm["A1"] = "SQA Testing Summary - Supporters Frame"
    sm["A1"].font = _font(bold=True, size=14, color="1F4E78")
    info = [("Website", WEBSITE), ("Prepared by", AUTHOR), ("Testing date", TEST_DATE),
            ("Environment", ENVIRONMENT), ("Total bugs reported", len(bugs))]
    r = 3
    for label, value in info:
        sm.cell(row=r, column=1, value=label).font = _font(bold=True)
        c = sm.cell(row=r, column=2, value=value)
        c.font = _font()
        c.alignment = Alignment(horizontal="left")
        r += 1
    r += 1
    sev_top = r
    r = _count_table(sm, r, "Bugs by Severity", SEVERITIES,
                     Counter(b["severity"] for b in bugs), colour_by_severity=True)
    r = _count_table(sm, r, "Bugs by Status", STATUSES, Counter(b["status"] for b in bugs))

    sm.cell(row=r, column=1, value="Lighthouse (homepage, mobile)").font = _font(bold=True, size=12, color="1F4E78")
    r += 1
    for k, v in LIGHTHOUSE.items():
        sm.cell(row=r, column=1, value=k).font = _font(bold=True)
        c = sm.cell(row=r, column=2, value=v)
        c.font = _font()
        c.alignment = Alignment(horizontal="left")
        r += 1
    r += 1
    sm.cell(row=r, column=1, value="Observations not counted as bugs").font = _font(bold=True, size=12, color="1F4E78")
    r += 1
    _header_row(sm, r, ["ID / Status", "Observation"])
    sm.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    r += 1
    for o in OBSERVATIONS:
        a = sm.cell(row=r, column=1, value=f"{o['id']}\n{o['status']}")
        b = sm.cell(row=r, column=2, value=f"{o['title']}. {o['note']}")
        for col in range(1, 8):
            sm.cell(row=r, column=col).border = BORDER
        for c in (a, b):
            c.font = _font()
            c.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
        sm.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        sm.row_dimensions[r].height = _estimate_height([(a.value, 30), (b.value, 88)])
        r += 1
    r += 1
    sm.cell(row=r, column=1, value="Current verification limitation").font = _font(
        bold=True, size=12, color="1F4E78")
    r += 1
    sm.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    note = sm.cell(row=r, column=1, value=COVERAGE_NOTE)
    note.font = _font(color="9C0006")
    note.alignment = Alignment(wrap_text=True, vertical="top")
    sm.row_dimensions[r].height = _estimate_height([(COVERAGE_NOTE, 100)])

    chart = BarChart()
    chart.type = "col"
    chart.title = "Bugs by Severity"
    chart.y_axis.title = "Number of bugs"
    chart.legend = None
    data = Reference(sm, min_col=2, min_row=sev_top + 1, max_row=sev_top + 1 + len(SEVERITIES))
    cats = Reference(sm, min_col=1, min_row=sev_top + 2, max_row=sev_top + 1 + len(SEVERITIES))
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.height, chart.width = 7.5, 11
    sm.add_chart(chart, f"D{sev_top}")

    # ---- Sheet 3: Test Log ----
    widths = [9, 18, 42, 48, 18]
    _header_row(tl, 1, ["Test ID", "Area", "Scenario", "Result", "Linked Bug"])
    for i, w in enumerate(widths, start=1):
        tl.column_dimensions[get_column_letter(i)].width = w
    for row, rec in enumerate(TEST_LOG, start=2):
        for col, value in enumerate(rec, start=1):
            c = tl.cell(row=row, column=col, value=value)
            c.font, c.border = _font(), BORDER
            c.alignment = Alignment(wrap_text=True, vertical="top",
                                    horizontal="center" if col in (1, 5) else "left")
        res = str(rec[3]).lower()
        if "not tested" in res or "not included" in res:
            tl.cell(row=row, column=4).fill = PatternFill("solid", fgColor="EDEDED")
        elif "passed" in res:
            tl.cell(row=row, column=4).fill = SEVERITY_FILL["Low"]
    tl.freeze_panes = "A2"
    tl.auto_filter.ref = f"A1:E{len(TEST_LOG) + 1}"

    # ---- Sheet 4: Coverage and limitations ----
    coverage.column_dimensions["A"].width = 28
    coverage.column_dimensions["B"].width = 16
    coverage.column_dimensions["C"].width = 62
    coverage.column_dimensions["D"].width = 68
    coverage.merge_cells("A1:D1")
    coverage["A1"] = "Test Coverage and Limitations"
    coverage["A1"].font = _font(bold=True, size=14, color="1F4E78")
    coverage.merge_cells("A2:D2")
    coverage["A2"] = COVERAGE_NOTE
    coverage["A2"].font = _font(color="9C0006")
    coverage["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    coverage.row_dimensions[2].height = _estimate_height([(COVERAGE_NOTE, 100)])
    _header_row(coverage, 4, ["Area / Page", "Coverage", "What was checked", "Not covered / limitation"])
    for row, item in enumerate(COVERAGE, start=5):
        for col, value in enumerate(item, start=1):
            cell = coverage.cell(row=row, column=col, value=value)
            cell.font, cell.border = _font(), BORDER
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        status = coverage.cell(row=row, column=2)
        if status.value == "Not tested":
            status.fill = PatternFill("solid", fgColor="EDEDED")
        elif status.value == "Blocked":
            status.fill = PatternFill("solid", fgColor="F8CBAD")
        else:
            status.fill = PatternFill("solid", fgColor="FFE699")
        status.font = _font(bold=True)
        coverage.row_dimensions[row].height = _estimate_height(
            [(item[0], 28), (item[1], 16), (item[2], 62), (item[3], 68)])
    coverage.freeze_panes = "A5"
    coverage.auto_filter.ref = f"A4:D{len(COVERAGE) + 4}"

    # ---- Sheet 5: Prioritized retest plan ----
    for column, width in enumerate((12, 10, 30, 34, 76, 76, 68), start=1):
        retest.column_dimensions[get_column_letter(column)].width = width
    retest.merge_cells("A1:G1")
    retest["A1"] = "Prioritized Retest Plan - Execute When the Site Is Available"
    retest["A1"].font = _font(bold=True, size=14, color="1F4E78")
    retest.merge_cells("A2:G2")
    retest["A2"] = COVERAGE_NOTE
    retest["A2"].font = _font(color="9C0006")
    retest["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    retest.row_dimensions[2].height = _estimate_height([(COVERAGE_NOTE, 105)])
    _header_row(retest, 4, [
        "Test ID", "Priority", "Test area", "Browser / viewport",
        "Procedure", "Pass criteria", "Execution status",
    ])
    for row, test in enumerate(RETEST_PLAN, start=5):
        values = [*test, RETEST_STATUS]
        for column, value in enumerate(values, start=1):
            cell = retest.cell(row=row, column=column, value=value)
            cell.font, cell.border = _font(), BORDER
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        priority = retest.cell(row=row, column=2)
        priority.fill = PatternFill(
            "solid", fgColor="F8CBAD" if priority.value == "P1" else "FFE699")
        priority.font = _font(bold=True)
        status = retest.cell(row=row, column=7)
        status.fill = PatternFill("solid", fgColor="EDEDED")
        status.font = _font(color="666666")
        retest.row_dimensions[row].height = _estimate_height([
            (values[0], 12), (values[1], 10), (values[2], 30), (values[3], 34),
            (values[4], 76), (values[5], 76), (values[6], 68),
        ])
    retest.freeze_panes = "A5"
    retest.auto_filter.ref = f"A4:G{len(RETEST_PLAN) + 4}"

    wb.save(path)
    return results


def export_to_available_path(path):
    """If the requested workbook is locked, save to a new sibling filename."""
    try:
        return path, export_to_excel(path)
    except PermissionError:
        directory, filename = os.path.split(path)
        name, extension = os.path.splitext(filename)
        index = 1
        while True:
            fallback = os.path.join(directory, f"{name}_{index}{extension}")
            if not os.path.exists(fallback):
                return fallback, export_to_excel(fallback)
            index += 1


# ======================================================================
# SECTION C - DESKTOP APPLICATION (Tkinter)
# ======================================================================

SEV_ROW_COLOURS = {"High": "#F8CBAD", "Medium": "#FFE699", "Low": "#C6E0B4"}
NAVY = "#1F4E78"
DEFAULT_XLSX = "Bug_Report_Sumaiya_Nasrin.xlsx"

if TK_OK:

    class BugReportApp(tk.Tk):
        def __init__(self):
            super().__init__()
            self.title(f"Supporters Frame - SQA Bug Report ({AUTHOR})")
            self.geometry("1180x780")
            self.minsize(980, 640)
            self.visible = []
            self.sev_var = tk.StringVar(value="All")
            self.sta_var = tk.StringVar(value="All")
            self.status_var = tk.StringVar()
            self._build_style()
            self._build_header()
            self._build_toolbar()
            self._build_statusbar()   # packed before the body so it is never squeezed out
            self._build_body()
            self.apply_filters()

        # ----- layout -----
        def _build_style(self):
            style = ttk.Style(self)
            try:
                style.theme_use("clam")
            except tk.TclError:
                pass
            style.configure("Treeview", rowheight=26, font=("Segoe UI", 10))
            style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"),
                            background=NAVY, foreground="white")
            style.map("Treeview.Heading", background=[("active", "#2E75B6")])
            style.configure("TButton", font=("Segoe UI", 10), padding=6)
            style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), padding=6)

        def _build_header(self):
            bar = tk.Frame(self, bg=NAVY)
            bar.pack(fill="x")
            tk.Label(bar, text="Supporters Frame - SQA Bug Report", bg=NAVY, fg="white",
                     font=("Segoe UI", 16, "bold")).pack(side="left", padx=14, pady=(10, 2))
            tk.Label(bar, text=f"{WEBSITE}   |   Prepared by {AUTHOR}   |   {TEST_DATE}",
                     bg=NAVY, fg="#DDEBF7", font=("Segoe UI", 9)).pack(
                side="left", padx=4, pady=(14, 2))

        def _build_toolbar(self):
            tb = ttk.Frame(self, padding=(12, 8))
            tb.pack(fill="x")
            ttk.Label(tb, text="Severity:").pack(side="left")
            sev = ttk.Combobox(tb, textvariable=self.sev_var, state="readonly", width=10,
                               values=["All"] + SEVERITIES)
            sev.pack(side="left", padx=(4, 12))
            ttk.Label(tb, text="Status:").pack(side="left")
            sta = ttk.Combobox(tb, textvariable=self.sta_var, state="readonly", width=12,
                               values=["All"] + STATUSES)
            sta.pack(side="left", padx=(4, 12))
            sev.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())
            sta.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())
            ttk.Button(tb, text="Reset filters", command=self.reset_filters).pack(side="left")
            ttk.Button(tb, text="Export to Excel", style="Accent.TButton",
                       command=self.export_excel).pack(side="right")
            ttk.Button(tb, text="Check evidence", command=self.check_evidence).pack(side="right", padx=6)
            ttk.Button(tb, text="Summary", command=self.show_summary).pack(side="right")

        def _build_body(self):
            pane = ttk.PanedWindow(self, orient="vertical")
            pane.pack(fill="both", expand=True, padx=12, pady=(0, 6))

            top = ttk.Frame(pane)
            cols = ("id", "title", "module", "severity", "status")
            self.tree = ttk.Treeview(top, columns=cols, show="headings", selectmode="browse", height=8)
            for col, text, width in (("id", "Bug ID", 80), ("title", "Bug Title / Summary", 560),
                                     ("module", "Module", 230), ("severity", "Severity", 80),
                                     ("status", "Status", 80)):
                self.tree.heading(col, text=text)
                self.tree.column(col, width=width, anchor="w" if col in ("title", "module") else "center")
            for sev, colour in SEV_ROW_COLOURS.items():
                self.tree.tag_configure(sev, background=colour)
            sb = ttk.Scrollbar(top, orient="vertical", command=self.tree.yview)
            self.tree.configure(yscrollcommand=sb.set)
            self.tree.pack(side="left", fill="both", expand=True)
            sb.pack(side="right", fill="y")
            self.tree.bind("<<TreeviewSelect>>", self.on_select)
            pane.add(top, weight=1)

            bottom = ttk.LabelFrame(pane, text="Bug details", padding=6)
            self.detail = tk.Text(bottom, wrap="word", font=("Segoe UI", 10), padx=10, pady=8,
                                  relief="flat", bg="#FAFAFA", state="disabled")
            dsb = ttk.Scrollbar(bottom, orient="vertical", command=self.detail.yview)
            self.detail.configure(yscrollcommand=dsb.set)
            self.detail.pack(side="left", fill="both", expand=True)
            dsb.pack(side="right", fill="y")
            self.detail.tag_configure("label", font=("Segoe UI", 10, "bold"), foreground=NAVY)
            self.detail.tag_configure("body", font=("Segoe UI", 10))
            self.detail.tag_configure("ok", foreground="#2E7D32")
            self.detail.tag_configure("bad", foreground="#C00000")
            for sev, colour in SEV_ROW_COLOURS.items():
                self.detail.tag_configure("sev_" + sev, background=colour, font=("Segoe UI", 10, "bold"))
            pane.add(bottom, weight=3)

        def _build_statusbar(self):
            ttk.Label(self, textvariable=self.status_var, relief="sunken", anchor="w",
                      padding=(8, 3)).pack(fill="x", side="bottom")

        # ----- behaviour -----
        def apply_filters(self):
            sev, sta = self.sev_var.get(), self.sta_var.get()
            self.visible = [b for b in BUGS
                            if (sev == "All" or b["severity"] == sev)
                            and (sta == "All" or b["status"] == sta)]
            self.tree.delete(*self.tree.get_children())
            for b in self.visible:
                self.tree.insert("", "end", iid=b["id"], tags=(b["severity"],),
                                 values=(b["id"], b["title"], b["module"], b["severity"], b["status"]))
            self.status_var.set(f"Showing {len(self.visible)} of {len(BUGS)} bugs")
            self._write_detail(None)

        def reset_filters(self):
            self.sev_var.set("All")
            self.sta_var.set("All")
            self.apply_filters()

        def on_select(self, _event=None):
            sel = self.tree.selection()
            if not sel:
                return
            bug = next((b for b in BUGS if b["id"] == sel[0]), None)
            self._write_detail(bug)

        def _write_detail(self, bug):
            t = self.detail
            t.configure(state="normal")
            t.delete("1.0", "end")
            if bug is None:
                t.insert("end", "Select a bug in the table above to see its full details.", "body")
            else:
                def field(label, text):
                    t.insert("end", label + "\n", "label")
                    t.insert("end", text + "\n\n", "body")
                field("Bug ID", bug["id"])
                field("Title", bug["title"])
                field("Module / Feature", bug["module"])
                field("Preconditions", bug["preconditions"])
                field("Steps to Reproduce", steps_as_text(bug))
                field("Expected Result", bug["expected"])
                field("Actual Result", bug["actual"])
                t.insert("end", "Severity\n", "label")
                t.insert("end", f" {bug['severity']} ", "sev_" + bug["severity"])
                t.insert("end", "\n\n", "body")
                field("Why this severity", bug["severity_reason"])
                field("Test Environment", ENVIRONMENT)
                t.insert("end", "Screenshot / Evidence\n", "label")
                for name in bug["evidence"]:
                    found = evidence_exists(name)
                    t.insert("end", f"  {name}  ", "body")
                    t.insert("end", "[found]\n" if found else "[NOT FOUND in evidence folder]\n",
                             "ok" if found else "bad")
                t.insert("end", "\n", "body")
                field("Reproducibility", bug["reproducibility"])
                field("Status", bug["status"])
            t.configure(state="disabled")
            t.yview_moveto(0)

        def show_summary(self):
            win = tk.Toplevel(self)
            win.title("Summary of Findings")
            win.geometry("620x640")
            win.transient(self)
            frame = ttk.Frame(win, padding=14)
            frame.pack(fill="both", expand=True)
            ttk.Label(frame, text="Summary of Findings", font=("Segoe UI", 15, "bold")).pack(anchor="w")
            ttk.Label(frame, text=f"{len(BUGS)} bugs reported on {WEBSITE}").pack(anchor="w", pady=(0, 8))

            ttk.Label(frame, text="By severity", font=("Segoe UI", 11, "bold")).pack(anchor="w")
            sev_counts = Counter(b["severity"] for b in BUGS)
            for sev in SEVERITIES:
                row = tk.Frame(frame)
                row.pack(fill="x", pady=1)
                tk.Label(row, text=f" {sev} ", bg=SEV_ROW_COLOURS[sev], width=10, anchor="w",
                         font=("Segoe UI", 10, "bold")).pack(side="left")
                tk.Label(row, text=str(sev_counts.get(sev, 0)), font=("Segoe UI", 10)).pack(side="left", padx=8)
            ttk.Label(frame, text="By status", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(10, 0))
            sta_counts = Counter(b["status"] for b in BUGS)
            for sta in STATUSES:
                ttk.Label(frame, text=f"   {sta}: {sta_counts.get(sta, 0)}").pack(anchor="w")

            ttk.Label(frame, text="Lighthouse (homepage, mobile)",
                      font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(10, 0))
            for k, v in LIGHTHOUSE.items():
                ttk.Label(frame, text=f"   {k}: {v}").pack(anchor="w")

            ttk.Label(frame, text="Observations not counted as bugs",
                      font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(10, 0))
            for o in OBSERVATIONS:
                ttk.Label(frame, text=f"   {o['id']} ({o['status']}): {o['title']}",
                          wraplength=560, justify="left").pack(anchor="w")
            ttk.Button(frame, text="Close", command=win.destroy).pack(anchor="e", pady=(12, 0))

        def check_evidence(self):
            miss = missing_evidence()
            folder = evidence_dir()
            if not os.path.isdir(folder):
                messagebox.showwarning("Evidence folder missing",
                                       f"The evidence folder does not exist:\n{folder}\n\n"
                                       "Create it and save your screenshots inside.")
                return
            if miss:
                messagebox.showwarning("Evidence check",
                                       "These screenshot files were NOT found:\n\n  - "
                                       + "\n  - ".join(miss) + f"\n\nFolder: {folder}")
            else:
                messagebox.showinfo("Evidence check", "All evidence files were found.")

        def export_excel(self):
            path = filedialog.asksaveasfilename(
                title="Save Excel bug report", defaultextension=".xlsx",
                initialdir=app_dir(), initialfile=DEFAULT_XLSX,
                filetypes=[("Excel workbook", "*.xlsx")])
            if not path:
                return
            try:
                saved_path, results = export_to_available_path(path)
            except PermissionError:
                messagebox.showerror("Export failed",
                                     "The report could not be saved. Close the workbook in Excel "
                                     "or choose a folder where you have write permission.")
                return
            except Exception as exc:
                messagebox.showerror("Export failed", f"Could not create the Excel file:\n{exc}")
                return
            embedded = sum(1 for v in results.values() if v == "embedded")
            problems = [n for n, v in results.items() if v != "embedded"]
            msg = f"Excel report saved:\n{saved_path}\n\nScreenshots embedded: {embedded} of {len(results)}"
            if saved_path != path:
                msg = ("The requested workbook could not be overwritten, so the report was saved "
                       f"as a new file.\n\n{msg}")
            if problems:
                msg += "\n\nNot embedded:\n  - " + "\n  - ".join(problems)
                messagebox.showwarning("Export finished with warnings", msg)
            else:
                messagebox.showinfo("Export complete", msg)


def main():
    if "--export" in sys.argv:
        idx = sys.argv.index("--export")
        out = sys.argv[idx + 1] if len(sys.argv) > idx + 1 else os.path.join(app_dir(), DEFAULT_XLSX)
        try:
            saved_path, results = export_to_available_path(out)
        except PermissionError as exc:
            print(f"Could not save the Excel file. Close it in Excel or choose a writable location.\n{exc}",
                  file=sys.stderr)
            return 1
        print("Excel file created:", saved_path)
        if saved_path != out:
            print("The requested file was locked; saved to a new filename instead.")
        bad = [n for n, v in results.items() if v != "embedded"]
        print(f"Screenshots embedded: {len(results) - len(bad)} of {len(results)}")
        if bad:
            print("Not embedded (missing or unreadable):")
            for n in bad:
                print("  -", n)
        return
    if not TK_OK:
        print("Tkinter is not available, so the window cannot open. "
              "Run with --export to create the Excel file only.")
        return
    try:
        BugReportApp().mainloop()
    except Exception as exc:  # last-resort message so the app never fails silently
        try:
            messagebox.showerror("Application error", f"The application stopped unexpectedly:\n{exc}")
        except Exception:
            print("Application error:", exc)


if __name__ == "__main__":
    sys.exit(main() or 0)
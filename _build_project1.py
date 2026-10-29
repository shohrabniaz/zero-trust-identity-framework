"""Build Project 1 — Zero-Trust Identity Framework submission pack."""
from __future__ import annotations

from datetime import date
from pathlib import Path
import json
import shutil

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from pptx import Presentation
from pptx.dml.color import RGBColor as PRgb
from pptx.util import Inches, Pt as PPt

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "Final_Submission"
AUDIT = ROOT / "01_Audit"
ARCH = ROOT / "02_Architecture"
ROAD = ROOT / "03_Roadmap"
EVAL = ROOT / "04_Evaluation"
EVID = ROOT / "evidence"
NAVY = RGBColor(0x1F, 0x4E, 0x79)
TODAY = date.today().isoformat()
FIRM = "Southern Ledger Partners (fictional regional accounting firm)"


def heading(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    for r in h.runs:
        r.font.color.rgb = NAVY


def write_synthetic_audit_json():
    EVID.mkdir(parents=True, exist_ok=True)
    data = {
        "Organisation": FIRM,
        "Mode": "Synthetic",
        "CollectedUtc": f"{TODAY}T00:00:00Z",
        "StaffModelled": 40,
        "Offices": ["Adelaide CBD", "Regional North", "Regional South"],
        "IdentityFindings": [
            {
                "Id": "ID-01",
                "Finding": "Shared legacy SaaS logins for practice-management style tools",
                "Risk": "Critical",
                "UsersAffected": 8,
                "Notes": "No individual accountability; password reuse likely",
            },
            {
                "Id": "ID-02",
                "Finding": "Standing Global Admin overlap with daily mailboxes",
                "Risk": "High",
                "UsersAffected": 3,
                "Notes": "No separate admin UPNs; no JIT",
            },
            {
                "Id": "ID-03",
                "Finding": "MFA inconsistent for remote staff",
                "Risk": "High",
                "UsersAffected": 18,
                "Notes": "Security defaults not verified; no CA",
            },
            {
                "Id": "ID-04",
                "Finding": "Guest accountants / contractors without expiry",
                "Risk": "Medium",
                "UsersAffected": 6,
                "Notes": "No access review cadence",
            },
            {
                "Id": "ID-05",
                "Finding": "Flat file-share access across offices",
                "Risk": "High",
                "UsersAffected": 40,
                "Notes": "Entitlements informal; leavers rarely removed promptly",
            },
            {
                "Id": "ID-06",
                "Finding": "Unmanaged BYOD accessing M365 from home",
                "Risk": "Medium",
                "UsersAffected": 22,
                "Notes": "No device compliance signal on Free tier",
            },
        ],
        "EntitlementHotspots": [
            "Tax working papers library — all seniors Full Control",
            "Payroll folder — two admins + one shared mailbox delegate",
            "Practice SaaS — shared 'office' account",
        ],
        "PrePilotRiskIndex": 72,
        "RiskIndexMethod": "Weighted count of Critical=25, High=15, Medium=8, Low=3 then scaled 0-100 (synthetic)",
    }
    path = EVID / f"{TODAY.replace('-', '')}_Identity_Entitlement_Audit.json"
    path.write_text(json.dumps(data, indent=2), encoding="utf-8")
    return data


def build_audit_doc(data):
    AUDIT.mkdir(parents=True, exist_ok=True)
    d = Document()
    d.add_heading("Identity and entitlement audit report", 0)
    d.add_paragraph(f"Md Shohrab Hossen Niaz | Project 1 | {TODAY}")
    d.add_paragraph(FIRM)
    heading(d, "Scope", 1)
    d.add_paragraph(
        "Synthetic current-state audit of ~40 identities across three offices. "
        "No live Graph export was used. Findings are labelled Synthetic and are "
        "designed to drive the zero-trust reference architecture."
    )
    heading(d, "Method", 1)
    d.add_paragraph(
        "Threat-informed checklist: privileged roles, MFA coverage, shared accounts, "
        "guests, file entitlements, device trust. Mapped to NIST SP 800-207 pillars "
        "(identity, device, network/application access) and ASD Essential Eight MFA/"
        "admin restriction themes."
    )
    heading(d, "Findings", 1)
    for f in data["IdentityFindings"]:
        d.add_paragraph(
            f"{f['Id']} [{f['Risk']}] {f['Finding']} — {f['Notes']} "
            f"(modelled users affected: {f['UsersAffected']})",
            style="List Bullet",
        )
    heading(d, "Entitlement hotspots", 1)
    for h in data["EntitlementHotspots"]:
        d.add_paragraph(h, style="List Number")
    heading(d, "Pre-pilot risk index", 1)
    d.add_paragraph(
        f"Synthetic pre-pilot identity risk index: {data['PrePilotRiskIndex']}/100. "
        f"Method: {data['RiskIndexMethod']}."
    )
    d.save(AUDIT / "Identity_Entitlement_Audit.docx")


def build_architecture():
    ARCH.mkdir(parents=True, exist_ok=True)
    d = Document()
    d.add_heading("Zero-trust identity reference architecture", 0)
    d.add_paragraph(f"Md Shohrab Hossen Niaz | Project 1 | {TODAY}")
    heading(d, "Design principle", 1)
    d.add_paragraph(
        "Identity is the primary perimeter. Every session is evaluated on user, "
        "device, location, and risk signals before access to mail, files, or "
        "accounting SaaS. Never trust the office LAN alone."
    )
    heading(d, "Logical architecture (text)", 1)
    for x in [
        "Users & guests → Entra ID (unique identities; no shared SaaS logins).",
        "Admin plane → separate admin UPNs + break-glass + time-boxed elevation.",
        "Policy engine → Phase 0: Security defaults / MFA; Phase 1: Conditional Access.",
        "Device signal → Phase 1: Intune compliance; Phase 0: educate + block hygiene.",
        "Apps → M365 + accounting SaaS via SSO/federation where available; block basic auth.",
        "Telemetry → sign-in logs + Secure Score tracking; ID Protection when licensed.",
    ]:
        d.add_paragraph(x, style="List Number")
    heading(d, "Trust policy set", 1)
    policies = [
        ("P0-MFA-ALL", "Phase 0", "Require MFA for all interactive users via security defaults or per-user MFA."),
        ("P0-ADMIN-SEP", "Phase 0", "Daily work on standard UPN; admin tasks on dedicated admin UPN."),
        ("P0-BREAKGLASS", "Phase 0", "Two monitored break-glass accounts excluded from MFA only with logging and vault."),
        ("P0-GUEST-TIGHT", "Phase 0", "Admins-only guest invite; 90-day guest disable if unused."),
        ("P0-SHARED-BAN", "Phase 0", "Retire shared SaaS passwords; one identity per person."),
        ("P1-CA-BASELINE", "Phase 1", "CA: require MFA for all users; block legacy authentication."),
        ("P1-CA-ADMIN", "Phase 1", "CA: admins require compliant or hybrid-joined device + MFA."),
        ("P1-CA-REMOTE", "Phase 1", "CA: external networks require MFA; optional compliant device for tax libraries."),
        ("P1-DEVICE", "Phase 1", "Intune compliance: disk encryption, patch level, no jailbreak."),
        ("P1-RISK", "Phase 1+", "Identity Protection / risk-based CA when P2 budget exists."),
    ]
    for pid, phase, text in policies:
        d.add_paragraph(f"{pid} ({phase}): {text}", style="List Bullet")
    heading(d, "Free-tier honesty", 1)
    d.add_paragraph(
        "Conditional Access, Identity Protection, and full Intune compliance are Phase 1+ "
        "purchases. Phase 0 still reduces risk by killing shared accounts, enforcing MFA, "
        "separating admin identities, and tightening guests — without inventing portal features."
    )
    d.save(ARCH / "ZeroTrust_Reference_Architecture_Policy_Set.docx")


def build_roadmap():
    ROAD.mkdir(parents=True, exist_ok=True)
    d = Document()
    d.add_heading("Phased implementation roadmap and cost estimates", 0)
    d.add_paragraph(f"Md Shohrab Hossen Niaz | Project 1 | {TODAY}")
    heading(d, "Phases", 1)
    phases = [
        (
            "Phase 0 — 0–6 weeks (one office pilot)",
            "Inventory identities; remove shared SaaS logins; enforce MFA; create admin UPNs; "
            "two break-glass accounts; guest invite lockdown; leaver checklist. Cost: staff time "
            "only (~40–60 engineer hours). Licence delta: $0 if M365 already owned.",
        ),
        (
            "Phase 1 — 6–16 weeks",
            "Move pilot office to Business Premium (or Entra ID P1 + Intune). Deploy baseline "
            "Conditional Access, block legacy auth, enrol devices, require compliance for admin "
            "and tax-library access. Indicative licence uplift: roughly AUD $10–25/user/month "
            "depending on current SKU (confirm with Microsoft price list at purchase time). "
            "Plus 80–120 hours integration/testing.",
        ),
        (
            "Phase 2 — 4–6 months",
            "Roll remaining offices; access reviews for guests; optional Identity Protection; "
            "SSO for accounting SaaS. Contingency 15% on licence and change-management time.",
        ),
    ]
    for title, body in phases:
        heading(d, title, 2)
        d.add_paragraph(body)
    heading(d, "Pilot scope (supervisor note)", 1)
    d.add_paragraph(
        "Pilot tightly to one office (~12–15 people) to stay trimester-achievable. "
        "Measure pre/post risk index, MFA coverage %, shared-account count, and user friction "
        "(helpdesk tickets in first two weeks)."
    )
    d.save(ROAD / "Phased_Roadmap_Cost_Estimates.docx")


def build_workbook(data):
    EVAL.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    fill = PatternFill("solid", fgColor="1F4E79")
    white = Font(color="FFFFFF", bold=True)

    ws = wb.active
    ws.title = "Pilot_Evaluation"
    ws.append(["Metric", "Pre-pilot (synthetic)", "Post-pilot target", "How measured"])
    for c in ws[1]:
        c.fill = fill
        c.font = white
    ws.append(["Identity risk index", data["PrePilotRiskIndex"], 35, "Weighted finding score"])
    ws.append(["MFA coverage %", 55, 100, "Registration / security defaults check"])
    ws.append(["Shared SaaS accounts", 3, 0, "App inventory + owner attestation"])
    ws.append(["Standing Global Admins (daily UPN)", 3, 0, "Directory roles + admin UPN split"])
    ws.append(["Guests without expiry review", 6, 0, "Guest export quarterly"])
    ws.append(["Helpdesk friction tickets (2 weeks)", "n/a", "<8", "Ticket tags MFA/CA"])
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 36

    find = wb.create_sheet("Audit_Findings")
    find.append(["Id", "Risk", "Finding", "UsersAffected", "Notes"])
    for c in find[1]:
        c.fill = fill
        c.font = white
    for f in data["IdentityFindings"]:
        find.append([f["Id"], f["Risk"], f["Finding"], f["UsersAffected"], f["Notes"]])
    for col in find.columns:
        find.column_dimensions[col[0].column_letter].width = 28
    find.column_dimensions["C"].width = 55

    cost = wb.create_sheet("Cost_Model")
    cost.append(["Phase", "Licence uplift (indicative)", "Effort hours", "Notes"])
    for c in cost[1]:
        c.fill = fill
        c.font = white
    cost.append(["0", "AUD 0", "40-60", "MFA, admin split, shared-account retirement"])
    cost.append(["1", "AUD 10-25 /user/month", "80-120", "Confirm SKU with Microsoft price list"])
    cost.append(["2", "As Phase 1 + optional P2", "60-100", "Rollout + reviews + SaaS SSO"])
    for col in cost.columns:
        cost.column_dimensions[col[0].column_letter].width = 30

    path = EVAL / "ZeroTrust_Pilot_Evaluation.xlsx"
    wb.save(path)
    return path


def build_eval_doc():
    d = Document()
    d.add_heading("Pilot evaluation and recommendations", 0)
    d.add_paragraph(f"Md Shohrab Hossen Niaz | Project 1 | {TODAY}")
    heading(d, "Pilot design", 1)
    d.add_paragraph(
        "One office, Phase 0 controls first, then optional Phase 1 CA if licences are funded. "
        "Pre/post metrics in the evaluation workbook. Collect short user feedback after week two."
    )
    heading(d, "Recommendations", 1)
    for x in [
        "Do not start with Identity Protection marketing — start with unique identities and MFA.",
        "Fund Business Premium (or equivalent) before promising Conditional Access to partners.",
        "Keep two break-glass accounts offline-documented and monitored.",
        "Tie leavers to the same-day removal from tax libraries and SaaS.",
        "Re-score the identity risk index after Phase 0 before buying P2.",
    ]:
        d.add_paragraph(x, style="List Number")
    d.save(EVAL / "Pilot_Evaluation_Recommendations.docx")


def build_report(data):
    sections = [
        (
            "Abstract",
            "I designed a zero-trust identity framework for a fictional regional accounting firm "
            "(Southern Ledger Partners, ~40 staff, three offices). The pack includes a synthetic "
            "identity and entitlement audit, a Free-tier-honest reference architecture and policy "
            "set, a phased roadmap with indicative costs, and a pilot evaluation workbook. Phase 0 "
            "reduces risk without Conditional Access. Phase 1 introduces CA and device compliance "
            "when licences are funded. Pre-pilot synthetic risk index is "
            f"{data['PrePilotRiskIndex']}/100 with a post-pilot target of 35.",
        ),
        (
            "1. Introduction",
            "Client trust and Tax Practitioners Board expectations make a breach existential for "
            "a regional accounting firm. The business context in the catalogue describes flat "
            "networking, shared legacy logins, informal access, and weak continuous verification. "
            "Project 1 asks which zero-trust identity controls are practical and cost-justified, "
            "how Conditional Access and device trust can be introduced without crushing a small "
            "non-technical workforce, and what measurable risk reduction a pilot can show.",
        ),
        (
            "2. Method",
            "Design science: synthetic current-state audit, threat-informed finding list, iterative "
            "reference architecture mapped to NIST SP 800-207 and ASD Essential Eight MFA/admin "
            "themes, phased roadmap, and pre/post pilot metrics. Live Entra Conditional Access "
            "was not configured because Free-tier laboratory access does not provide CA; the "
            "architecture states that limit instead of inventing portal evidence.",
        ),
        (
            "3. Current-state findings",
            "Critical and High issues dominate: shared SaaS logins, daily Global Admin use, "
            "inconsistent MFA, flat file entitlements, and unmanaged BYOD. Guests lack expiry "
            "reviews. These are process and identity-hygiene failures first — licence gaps second.",
        ),
        (
            "4. Reference architecture",
            "Unique identities feed Entra ID. A policy engine starts as Security defaults / MFA "
            "(Phase 0) and grows into Conditional Access (Phase 1). Admin plane uses separate "
            "UPNs, break-glass, and time-boxed elevation. Device compliance arrives with Intune "
            "in Phase 1. Accounting SaaS moves to SSO where possible. Telemetry is sign-in logs "
            "and Secure Score; Identity Protection is Phase 1+ only.",
        ),
        (
            "5. Roadmap and cost",
            "Phase 0 is staff time only. Phase 1 indicative licence uplift is about AUD $10–25 "
            "per user per month depending on current SKU, plus integration effort. Pilot one "
            "office first. Costs are indicative and must be confirmed against Microsoft pricing "
            "at purchase time — they are not quotes.",
        ),
        (
            "6. Evaluation",
            "Success is MFA coverage at 100% of interactive users, zero shared SaaS accounts, "
            "zero daily-UPN Global Admins, guests on a review cycle, risk index ≤35, and "
            "contained helpdesk friction. User feedback gates aggressive CA grants.",
        ),
        (
            "7. Conclusion",
            "Zero trust for this SME starts with identity hygiene and MFA, not a shopping list "
            "of every Entra add-on. A phased, Free-tier-honest design lets the firm show partners "
            "and insurers a credible plan while funding Conditional Access for the office that "
            "handles the most sensitive tax data first.",
        ),
        (
            "References",
            "NIST SP 800-207 Zero Trust Architecture. ASD Essential Eight. ISO/IEC 27001:2022 "
            "A.5.15–A.5.18. Privacy Act 1988 (APPs). Microsoft Learn: Conditional Access; "
            "security defaults; Identity Protection overview.",
        ),
    ]
    bw = sum(len(b.split()) for t, b in sections if t != "References")
    m = Document()
    tp = m.add_paragraph()
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = tp.add_run(
        "Designing a zero-trust identity framework\n"
        "for a regional accounting firm"
    )
    tr.bold = True
    tr.font.size = Pt(16)
    tr.font.color.rgb = NAVY
    mp = m.add_paragraph()
    mp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    mr = mp.add_run(
        f"Md Shohrab Hossen Niaz\nProject 1 — Zero-Trust Identity Framework\n"
        f"{FIRM}\nSynthetic audit | Free-tier-honest architecture\n"
        f"Word count (excluding title block and reference list): {bw}"
    )
    mr.font.size = Pt(11)
    for title, body in sections:
        heading(m, title, 1)
        m.add_paragraph(body)
    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "01_Project1_Report.docx"
    m.save(path)
    print("BODY_WORDS", bw)
    return path


def build_slides(data):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    navy = PRgb(0x1F, 0x4E, 0x79)

    def add(title, bullets, notes):
        s = prs.slides.add_slide(prs.slide_layouts[6])
        b = s.shapes.add_textbox(Inches(0.6), Inches(0.4), Inches(12), Inches(1))
        r = b.text_frame.paragraphs[0].add_run()
        r.text = title
        r.font.size = PPt(28)
        r.font.bold = True
        r.font.color.rgb = navy
        body = s.shapes.add_textbox(Inches(0.6), Inches(1.5), Inches(12), Inches(5))
        tf = body.text_frame
        tf.word_wrap = True
        first = True
        for line in bullets:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            run = p.add_run()
            run.text = line
            run.font.size = PPt(18)
            p.space_after = PPt(10)
        s.notes_slide.notes_text_frame.text = notes

    add(
        "Zero-trust identity for an accounting SME",
        ["Southern Ledger Partners — fictional ~40 staff", "Identity as the perimeter", "Phase 0 hygiene → Phase 1 Conditional Access"],
        "Project 1.",
    )
    add(
        "The problem",
        ["Shared SaaS logins and standing admins", "Inconsistent MFA; flat file access", "Cannot show continuous verification"],
        "Why ZT identity.",
    )
    add(
        "The design",
        ["Unique IDs + MFA + admin split (Phase 0)", "CA + Intune compliance when funded (Phase 1)", f"Risk index {data['PrePilotRiskIndex']} → target 35"],
        "Architecture.",
    )
    add(
        "Pilot and ask",
        ["One office first; measure MFA, shared accounts, friction", "Costs are indicative, confirm Microsoft pricing", "No invented CA screenshots on Free tier"],
        "Questions?",
    )
    path = OUT / "02_Project1_Presentation.pptx"
    prs.save(path)
    return path


def assemble():
    data = write_synthetic_audit_json()
    build_audit_doc(data)
    build_architecture()
    build_roadmap()
    build_workbook(data)
    build_eval_doc()
    build_report(data)
    build_slides(data)

    mapping = [
        (EVAL / "ZeroTrust_Pilot_Evaluation.xlsx", OUT / "Appendix" / "A_Evaluation" / "ZeroTrust_Pilot_Evaluation.xlsx"),
        (EVAL / "Pilot_Evaluation_Recommendations.docx", OUT / "Appendix" / "A_Evaluation" / "Pilot_Evaluation_Recommendations.docx"),
        (AUDIT / "Identity_Entitlement_Audit.docx", OUT / "Appendix" / "B_Audit" / "Identity_Entitlement_Audit.docx"),
        (ARCH / "ZeroTrust_Reference_Architecture_Policy_Set.docx", OUT / "Appendix" / "C_Architecture" / "ZeroTrust_Reference_Architecture_Policy_Set.docx"),
        (ROAD / "Phased_Roadmap_Cost_Estimates.docx", OUT / "Appendix" / "D_Roadmap" / "Phased_Roadmap_Cost_Estimates.docx"),
    ]
    for src, dst in mapping:
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)
    app_e = OUT / "Appendix" / "E_Evidence"
    app_e.mkdir(parents=True, exist_ok=True)
    for f in EVID.glob("*.json"):
        shutil.copy2(f, app_e / f.name)

    readme = Document()
    readme.add_heading("Project submission", 0)
    readme.add_paragraph(
        "Md Shohrab Hossen Niaz\nProject 1 — Zero-Trust Identity Framework\n"
        "Southern Ledger Partners (fictional). Synthetic audit; Free-tier-honest design."
    )
    readme.save(OUT / "00_Read_Me_First.docx")
    print("assembled", OUT)


if __name__ == "__main__":
    assemble()

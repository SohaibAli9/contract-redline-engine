"""Generate INVENTED construction subcontracts for the demo.

Every company, project and person here is made up. The clause wording is written for this demo
in the style of a general contractor's form subcontract; it is not copied from AIA, ConsensusDocs
or any other copyrighted form.

Output goes to DATA_DIR (default ./data):
  historical/   10 contracts already marked up by the (invented) subcontractor's reviewer:
                8 .docx with Word Track Changes + comments, 2 .pdf with strikeouts + popup notes
  incoming/     5 held-out contracts, clean: 3 .docx, 2 .pdf
  key.json      the redlines a human reviewer would make on each incoming contract, written
                here, before any model run
"""
from __future__ import annotations

import json
import os
import random
import re
import subprocess
from pathlib import Path

import pymupdf
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

from redline.writers.docx_writer import DocxEdit, apply_edits

DATA = Path(os.environ.get("DATA_DIR", "data"))
SUB = "Bayview Mechanical LLC"
GCS = ["Northgate Builders, Inc.", "Calder & Price Construction Co.", "Summit Ridge Contractors, LLC",
       "Ironline Constructors, Inc.", "Westbrook General Contracting", "Harlow Kemp Builders",
       "Pinecrest Construction Group", "Ostrander Brothers, Inc.", "Meridian Civil & Building",
       "Talbot Yard Construction", "Redfern Hollis Contractors", "Granite Point Builders",
       "Lakeshore Design-Build, LLC", "Fairmount Construction Services", "Carver Lane Contracting"]
PROJECTS = ["Riverside Medical Office Building", "Elm Street Elementary School Addition",
            "Harbor Point Apartments, Phase 2", "County Public Safety Building",
            "Westgate Distribution Center", "Oak Hollow Senior Living", "Main & 5th Mixed-Use",
            "Kingsley Library Renovation", "North Campus Parking Structure", "Mill Creek Lab Fit-Out",
            "Ashford Fire Station No. 4", "Cedar Ridge Middle School", "Bayside Hotel Renovation",
            "Lincoln Avenue Transit Center", "Stonebridge Data Hall"]
REVIEWERS = ["M. Reyes (Contracts)", "D. Okafor (Risk)", "M. Reyes (Contracts)", "S. Lindqvist (PM)"]

WORDS = {1: "one", 2: "two", 3: "three", 5: "five", 7: "seven", 10: "ten", 14: "fourteen",
         15: "fifteen", 21: "twenty-one", 30: "thirty", 45: "forty-five"}


def n(x: int) -> str:
    return f"{WORDS[x]} ({x})"


# --------------------------------------------------------------------------------------------
# Clause library. Each clause: number, heading, text (with **bold** runs), and the redlines the
# subcontractor's reviewer makes: (rule_id, original substring, replacement, comment).
# Variants keep wording different between contracts so the rules have to generalise.
# --------------------------------------------------------------------------------------------

def clauses(rng: random.Random, gc: str, project: str, v: dict) -> list[dict]:
    C = []

    def add(num, heading, text, edits=()):
        C.append({"num": num, "heading": heading, "text": text, "edits": list(edits)})

    add("1", "Subcontract Documents",
        f"This Subcontract between **{gc}** (\"Contractor\") and **{SUB}** (\"Subcontractor\") for the "
        f"{project} (the \"Project\") consists of this Agreement, the Prime Contract between Contractor and "
        "the Owner, the Drawings and Specifications listed in Exhibit A, and all addenda issued before execution. "
        "Subcontractor is bound to Contractor to the same extent Contractor is bound to the Owner under the "
        "Prime Contract, as far as the Prime Contract applies to the Work.")
    add("2", "Scope of Work",
        "Subcontractor shall furnish all labor, materials, equipment, supervision and services necessary to "
        "complete the plumbing, HVAC and controls work described in Exhibit B (the \"Work\"), in a good and "
        "workmanlike manner and in accordance with the Subcontract Documents.")
    add("3", "Subcontract Sum",
        f"Contractor shall pay Subcontractor the sum of ${v['sum']:,} (the \"Subcontract Sum\"), subject to "
        "additions and deductions by approved Change Order.")

    # 4. Payment: pay-if-paid
    if v["pay"] == "if":
        pay_text = rng.choice([
            "Receipt of payment by Contractor from the Owner for the Work is a condition precedent to "
            "Contractor's obligation to make any payment to Subcontractor.",
            "Payment by the Owner to Contractor is an express condition precedent to any payment by "
            "Contractor to Subcontractor, and Subcontractor assumes the risk of the Owner's nonpayment.",
        ])
        edits = [("PAY-1", pay_text,
                  "Payment by the Owner is not a condition precedent to payment to Subcontractor. Contractor "
                  "shall pay Subcontractor no later than forty-five (45) days after receipt of Subcontractor's "
                  "approved application, unless nonpayment results from Subcontractor's default.",
                  "Pay-if-paid is not accepted. Pay-when-paid with a 45 day outside date.")]
    else:
        pay_text = ("Contractor shall pay Subcontractor no later than forty-five (45) days after receipt of "
                    "Subcontractor's approved application, whether or not Contractor has been paid by the Owner.")
        edits = []
    add("4", "Progress Payments",
        f"Subcontractor shall submit applications for payment by the {rng.choice(['20th', '25th'])} day of each "
        f"month. Contractor shall pay Subcontractor within {n(rng.choice([7, 10]))} days after Contractor "
        f"receives payment from the Owner for the Work covered by the application. {pay_text}", edits)

    # 5. Retainage
    ret = v["ret"]
    ret_words = {10: "ten percent (10%)", 5: "five percent (5%)"}[ret]
    release = rng.choice(["final completion and acceptance of the entire Project by the Owner",
                          "final completion of the Project and final payment by the Owner to Contractor"])
    ret_text = (f"Contractor shall withhold retainage of {ret_words} from each progress payment until "
                f"{release}.")
    edits = []
    if ret == 10:
        edits.append(("RET-1", ret_words, "five percent (5%)",
                      "Retainage capped at 5%."))
    edits.append(("RET-2", release, "Substantial Completion of the Work",
                  "Retainage released at Substantial Completion of our Work, not the whole Project."))
    add("5", "Retainage", ret_text, edits)

    # 6. Liquidated damages
    ld = v["ld"]
    ld_text = (f"If the Work is not completed within the Subcontract Time, Subcontractor shall pay Contractor "
               f"liquidated damages of ${ld:,} per calendar day for each day of delay"
               + ("" if v["ld_capped"] else rng.choice([", without limitation.", ", without any cap or limit."]))
               + ("" if not v["ld_capped"] else ", to the extent the Owner assesses liquidated damages against "
                  "Contractor caused by Subcontractor's delay, not to exceed ten percent (10%) of the Subcontract Sum."))
    edits = []
    if not v["ld_capped"]:
        tail = ld_text[ld_text.index("for each day of delay"):]
        edits.append(("LD-1", tail,
                      "for each day of delay, only to the extent the Owner assesses liquidated damages against "
                      "Contractor caused by Subcontractor's delay, not to exceed ten percent (10%) of the "
                      "Subcontract Sum.",
                      "LDs are pass-through only, in proportion to our delay, capped at 10%."))
    add("6", "Liquidated Damages", ld_text, edits)

    # 7. Indemnity
    if v["indem"] == "broad":
        broad = rng.choice([
            "regardless of whether such claim is caused in part by a party indemnified hereunder",
            "whether or not caused in whole or in part by the negligence of Contractor or the Owner",
        ])
        edits = [("IND-1", broad,
                  "but only to the extent caused by the negligent acts or omissions of Subcontractor or anyone "
                  "for whom Subcontractor is responsible",
                  "Broad form indemnity not accepted. Limit to our share of fault.")]
    else:
        broad = ("but only to the extent caused by the negligent acts or omissions of Subcontractor or anyone "
                 "for whom Subcontractor is responsible")
        edits = []
    add("7", "Indemnification",
        "To the fullest extent permitted by law, Subcontractor shall defend, indemnify and hold harmless "
        "**Contractor**, the Owner and their officers, agents and employees from all claims, damages, losses "
        f"and expenses, including attorneys' fees, arising out of or resulting from the Work, {broad}.", edits)

    add("8", "Insurance",
        "Before starting the Work, Subcontractor shall furnish certificates of commercial general liability "
        "insurance with limits of not less than $1,000,000 per occurrence and $2,000,000 aggregate, "
        "workers' compensation at statutory limits, and automobile liability of $1,000,000, naming "
        "Contractor and the Owner as additional insureds.")

    # 9. Delay
    if v["nodelay"]:
        nd = rng.choice([
            "Subcontractor's sole remedy for any delay shall be an extension of the Subcontract Time, and "
            "Subcontractor shall not be entitled to any damages or additional compensation for delay.",
            "An extension of time is Subcontractor's exclusive remedy for delay of any kind, and no damages for "
            "delay shall be recoverable by Subcontractor.",
        ])
        edits = [("DEL-1", "", " This Section does not bar Subcontractor's recovery of delay damages to the extent "
                  "Contractor recovers such damages from the Owner on Subcontractor's behalf, or to the extent "
                  "the delay is caused by Contractor.",
                  "No-damages-for-delay needs a carve-out for Contractor-caused and Owner-paid delay.")]
    else:
        nd = ("Subcontractor shall be entitled to an extension of time and to recover its direct costs for delays "
              "caused by Contractor.")
        edits = []
    add("9", "Delays", f"{nd}", edits)

    # 10. Notice of claims
    nd_days = v["notice"]
    notice_phrase = f"within {n(nd_days)} days"
    edits = []
    if nd_days < 14:
        edits.append(("NOT-1", notice_phrase, "within fourteen (14) days",
                      "Claim notice of less than 14 days is not workable."))
    add("10", "Claims",
        f"Subcontractor shall give Contractor written notice of any claim {notice_phrase} after the event "
        "giving rise to the claim. Claims not made within that time are waived.", edits)

    # 11. Consequential damages
    if v["consq"] == "one-sided":
        cd = ("Subcontractor waives all claims against Contractor and the Owner for consequential damages "
              "arising out of or relating to this Subcontract.")
        edits = [("CONS-1", "Subcontractor waives all claims against Contractor and the Owner",
                  "Contractor and Subcontractor mutually waive all claims against each other",
                  "Waiver of consequential damages must be mutual.")]
    else:
        cd = ("Contractor and Subcontractor mutually waive all claims against each other for consequential "
              "damages arising out of or relating to this Subcontract.")
        edits = []
    add("11", "Consequential Damages", cd, edits)

    # 12. Warranty
    w = v["warranty"]
    w_phrase = f"for a period of {n(w)} {'year' if w == 1 else 'years'}"
    edits = []
    if w > 1:
        edits.append(("WAR-1", w_phrase, "for a period of one (1) year",
                      "Warranty period is one year from Substantial Completion."))
    add("12", "Warranty",
        f"Subcontractor warrants the Work against defects in materials and workmanship {w_phrase} after "
        "Substantial Completion of the Project, and shall promptly correct defective Work at its own cost.", edits)

    add("13", "Termination for Convenience",
        "Contractor may terminate this Subcontract for convenience on seven (7) days' written notice, in which "
        "case Subcontractor shall be paid for Work properly performed and reasonable demobilization costs.")
    add("14", "Dispute Resolution",
        f"Disputes shall be resolved by litigation in the state courts of the county in which the Project is "
        f"located. This Subcontract is governed by the law of the state in which the Project is located.")
    return C


def variant(rng: random.Random, force: dict | None = None) -> dict:
    v = {
        "sum": rng.choice([184_500, 412_750, 968_000, 1_245_300, 2_310_000, 655_420]),
        "pay": rng.choice(["if", "if", "if", "when"]),
        "ret": rng.choice([10, 10, 10, 5]),
        "ld": rng.choice([500, 750, 1_000, 1_500, 2_500]),
        "ld_capped": rng.random() < 0.2,
        "indem": rng.choice(["broad", "broad", "broad", "limited"]),
        "nodelay": rng.random() < 0.75,
        "notice": rng.choice([3, 5, 7, 7, 14, 21]),
        "consq": rng.choice(["one-sided", "one-sided", "mutual"]),
        "warranty": rng.choice([1, 2, 2, 3]),
    }
    v.update(force or {})
    return v


# --------------------------------------------------------------------------------------------
# DOCX building
# --------------------------------------------------------------------------------------------

def _add_runs(par, text: str, size=10.5):
    for i, piece in enumerate(re.split(r"\*\*", text)):
        if not piece:
            continue
        r = par.add_run(piece)
        r.font.size = Pt(size)
        r.bold = bool(i % 2)


def _page_field(par):
    for kind, txt in (("PAGE", "1"),):
        r = par.add_run()
        b = OxmlElement("w:fldChar"); b.set(qn("w:fldCharType"), "begin"); r._r.append(b)
        r2 = par.add_run()
        it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = f" {kind} "; r2._r.append(it)
        r3 = par.add_run()
        s = OxmlElement("w:fldChar"); s.set(qn("w:fldCharType"), "separate"); r3._r.append(s)
        par.add_run(txt)
        r4 = par.add_run()
        e = OxmlElement("w:fldChar"); e.set(qn("w:fldCharType"), "end"); r4._r.append(e)


def build_docx(path: Path, gc: str, project: str, cl: list[dict], number_style: str, sub_no: str):
    doc = Document()
    sec = doc.sections[0]
    hp = sec.header.paragraphs[0]
    hp.text = f"Subcontract No. {sub_no}  |  {gc}  |  {project}"
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp.add_run("Subcontractor initials: ______    Page ")
    _page_field(fp)

    t = doc.add_paragraph(); t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run("STANDARD FORM SUBCONTRACT AGREEMENT"); r.bold = True; r.font.size = Pt(14)
    doc.add_paragraph(f"Subcontract No. {sub_no}. Invented sample for a software demonstration; "
                      "all parties and projects are fictitious.").runs[0].italic = True
    para_index = {}
    for c in cl:
        if number_style == "auto":
            # Word auto-numbering: the clause number is NOT in the paragraph text
            hp_ = doc.add_paragraph(style="List Number")
            hr = hp_.add_run(c["heading"]); hr.bold = True
        else:
            hp_ = doc.add_paragraph()
            hr = hp_.add_run(f"{c['num']}. {c['heading']}"); hr.bold = True
        bp = doc.add_paragraph()
        _add_runs(bp, c["text"])
        bp.paragraph_format.space_after = Pt(8)
        para_index[c["num"]] = len(doc.paragraphs) - 1
    doc.add_paragraph("")
    doc.add_paragraph(f"CONTRACTOR: {gc}            SUBCONTRACTOR: {SUB}")
    doc.save(path)
    return para_index


def plain(text: str) -> str:
    return text.replace("**", "")


def to_pdf(docx: Path, outdir: Path) -> Path:
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(outdir), str(docx)],
                   check=True, capture_output=True, timeout=120)
    return outdir / (docx.stem + ".pdf")


def mark_up_docx(path: Path, cl: list[dict], para_index: dict, rng: random.Random, author: str, date: str):
    edits = []
    for c in cl:
        text = plain(c["text"])
        for (_rule, orig, repl, comment) in c["edits"]:
            if rng.random() < 0.12:      # reviewers don't apply every rule every time
                continue
            if orig:
                s = text.index(orig); e = s + len(orig)
            else:
                s = e = len(text)
            edits.append(DocxEdit(para_index[c["num"]], s, e, repl, comment if rng.random() < 0.85 else ""))
    log = apply_edits(str(path), str(path), edits, author=author, date=date)
    assert not log, log


def mark_up_pdf(pdf: Path, cl: list[dict], rng: random.Random, author: str):
    doc = pymupdf.open(pdf)
    for c in cl:
        for (_rule, orig, repl, comment) in c["edits"]:
            if not orig or rng.random() < 0.12:
                continue
            for page in doc:
                quads = page.search_for(orig, quads=True)
                if quads:
                    a = page.add_strikeout_annot(quads)
                    a.set_info(content=f"Replace with: {repl}\n{comment}", title=author)
                    a.update()
                    break
    doc.save(pdf, incremental=True, encryption=pymupdf.PDF_ENCRYPT_KEEP)


def main():
    rng = random.Random(20261120)
    for d in ("historical", "incoming", "_tmp"):
        (DATA / d).mkdir(parents=True, exist_ok=True)
    key = {}
    plan = [("historical", f"H{i:02d}", "docx") for i in range(1, 9)] + \
           [("historical", f"H{i:02d}", "pdf") for i in (9, 10)] + \
           [("incoming", f"N{i:02d}", "docx") for i in (1, 2, 3)] + \
           [("incoming", f"N{i:02d}", "pdf") for i in (4, 5)]
    forced = {"N02": {"pay": "when", "ret": 5, "indem": "limited", "consq": "mutual", "warranty": 1,
                      "notice": 21, "nodelay": False, "ld_capped": True}}     # a clean contract: should get 0 edits
    for i, (kind, cid, fmt) in enumerate(plan):
        gc, project = GCS[i], PROJECTS[i]
        v = variant(rng, forced.get(cid))
        cl = clauses(rng, gc, project, v)
        number_style = "auto" if cid in ("H03", "H06", "N03") else "typed"
        stem = f"{cid}_{re.sub(r'[^A-Za-z]+', '_', gc).strip('_')}_Subcontract"
        out_docx = (DATA / "_tmp" if fmt == "pdf" else DATA / kind) / f"{stem}.docx"
        pidx = build_docx(out_docx, gc, project, cl, number_style, f"{24 + i:02d}-{1000 + i * 37}")
        if kind == "historical":
            author = REVIEWERS[i % len(REVIEWERS)]
            date = f"202{3 + i % 3}-0{1 + i % 9}-1{i % 9}T10:00:00Z"
            if fmt == "docx":
                mark_up_docx(out_docx, cl, pidx, rng, author, date)
            elif cid == "H09":
                # reviewer marked up a PDF with strikeout annotations and popup notes
                pdf = to_pdf(out_docx, DATA / kind)
                mark_up_pdf(pdf, cl, rng, author)
            else:
                # a printed redline: tracked changes exported to PDF, no annotations, only colour and lines
                mark_up_docx(out_docx, cl, pidx, rng, author, date)
                to_pdf(out_docx, DATA / kind)
        else:
            if fmt == "pdf":
                to_pdf(out_docx, DATA / kind)
            key[stem] = [{"clause": c["num"], "heading": c["heading"], "rule": e[0], "original": e[1],
                          "replacement": e[2]} for c in cl for e in c["edits"]]
    (DATA / "key.json").write_text(json.dumps(key, indent=1))
    print(json.dumps({k: len(v) for k, v in key.items()}))


if __name__ == "__main__":
    main()

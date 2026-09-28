# -*- coding: utf-8 -*-
"""Build 'Recognition Is Not Evidence' - a critical reading of Thierry
Ehrmann's 'Dialogue Between a Thinker and AI' (Raw Typescript, English
edition, August 2026).

The book under analysis is Aible Register entry 9. It is distributed by its
author under CC BY-NC-ND 4.0, which permits faithful quotation with full
attribution. Every quotation here is reproduced without modification and
carries its page number in the Raw Typescript PDF edition.

Ehrmann invites this: "Read it. Challenge it. Cite it. Share it." (p. 2)

Usage:  python build_critique.py <out.docx>
"""
import sys
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

SERIF = "Cambria"
MONO = "Consolas"
SYMBOL = "Times New Roman"
A4_W, A4_H = 8.268, 11.693
MARGIN_LR = 1.35
TEXT_W = A4_W - 2 * MARGIN_LR

doc = Document()
normal = doc.styles["Normal"]
normal.font.name = SERIF
normal.font.size = Pt(11.5)
normal._element.rPr.rFonts.set(qn("w:eastAsia"), SERIF)
normal.paragraph_format.space_after = Pt(8)
normal.paragraph_format.line_spacing = 1.25

sec = doc.sections[0]
sec.page_width = Inches(A4_W)
sec.page_height = Inches(A4_H)
sec.top_margin = Inches(1.2)
sec.bottom_margin = Inches(1.2)
sec.left_margin = Inches(MARGIN_LR)
sec.right_margin = Inches(MARGIN_LR)
sec.different_first_page_header_footer = False

_pending_break = [False]


def _consume(pf):
    if _pending_break[0]:
        pf.page_break_before = True
        _pending_break[0] = False


def _track(run, tw):
    rPr = run._element.get_or_add_rPr()
    el = OxmlElement("w:spacing")
    el.set(qn("w:val"), str(tw))
    rPr.append(el)


def para(text="", *, align=None, size=11.5, bold=False, italic=False, caps=False,
         font=SERIF, color=None, space_before=None, space_after=8, line=1.25,
         left=None, right=None, hanging=None, track=None, style=None,
         container=None, keep_next=False):
    p = (container.add_paragraph(style=style) if container is not None
         else doc.add_paragraph(style=style))
    pf = p.paragraph_format
    if align is not None:
        p.alignment = align
    pf.space_after = Pt(space_after)
    if space_before is not None:
        pf.space_before = Pt(space_before)
    pf.line_spacing = line
    pf.keep_with_next = keep_next
    if container is None:
        _consume(pf)
    if left is not None:
        pf.left_indent = Inches(left)
    if right is not None:
        pf.right_indent = Inches(right)
    if hanging is not None:
        pf.first_line_indent = Inches(-hanging)
    if text:
        for si, seg in enumerate(text.split("|")):
            if seg == "":
                continue
            f = SYMBOL if si % 2 == 1 else font
            for ii, chunk in enumerate(seg.split("*")):
                if chunk == "":
                    continue
                r = p.add_run(chunk)
                r.font.name = f
                rPr = r._element.get_or_add_rPr()
                for a in ("w:ascii", "w:hAnsi", "w:cs"):
                    rPr.rFonts.set(qn(a), f)
                r.font.size = Pt(size)
                r.bold = bold
                r.italic = (not italic) if (ii % 2 == 1) else italic
                r.font.all_caps = caps
                if color:
                    r.font.color.rgb = RGBColor.from_string(color)
                if track:
                    _track(r, track)
    return p


def blank(pts=10):
    return para("", space_after=pts)


def page_break():
    _pending_break[0] = True


def chapter(title, caps=True):
    p = doc.add_paragraph(style="Heading 1")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(24)
    p.paragraph_format.space_after = Pt(14)
    p.paragraph_format.keep_with_next = True
    _consume(p.paragraph_format)
    r = p.add_run(title)
    r.font.name = SERIF
    r.font.size = Pt(16)
    r.bold = False
    r.font.all_caps = caps
    r.font.color.rgb = RGBColor.from_string("1A1A1A")
    _track(r, 60)
    return p


def rule_line():
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_after = Pt(20)
    pf.left_indent = Inches(1.6)
    pf.right_indent = Inches(1.6)
    pf.keep_with_next = True
    pPr = p._element.get_or_add_pPr()
    bd = OxmlElement("w:pBdr")
    b = OxmlElement("w:bottom")
    b.set(qn("w:val"), "single"); b.set(qn("w:sz"), "4")
    b.set(qn("w:space"), "1"); b.set(qn("w:color"), "999999")
    bd.append(b); pPr.append(bd)
    return p


def code(text, container=None):
    return para(text, font=MONO, size=10.5, color="333333",
                space_after=3, line=1.15, left=1.35, container=container)


def body(text, container=None):
    return para(text, align=WD_ALIGN_PARAGRAPH.JUSTIFY, container=container)


def standout(text, container=None):
    return para(text, align=WD_ALIGN_PARAGRAPH.CENTER, italic=True,
                space_before=8, space_after=12, container=container)


def cited(text):
    """A small grey provenance line under a claim."""
    return para(text, align=WD_ALIGN_PARAGRAPH.LEFT, size=9.5, color="777777",
                space_before=2, space_after=12, left=0.35, line=1.15)


def box_cell(width_in):
    t = doc.add_table(rows=1, cols=1)
    t.autofit = False
    borders = OxmlElement("w:tblBorders")
    for s in ("top", "left", "bottom", "right"):
        el = OxmlElement("w:" + s)
        el.set(qn("w:val"), "single"); el.set(qn("w:sz"), "6")
        el.set(qn("w:space"), "0"); el.set(qn("w:color"), "000000")
        borders.append(el)
    t._tbl.tblPr.append(borders)
    mar = OxmlElement("w:tblCellMar")
    for s, v in (("top", 220), ("left", 260), ("bottom", 220), ("right", 260)):
        el = OxmlElement("w:" + s)
        el.set(qn("w:w"), str(v)); el.set(qn("w:type"), "dxa")
        mar.append(el)
    t._tbl.tblPr.append(mar)
    c = t.rows[0].cells[0]
    c.width = Inches(width_in)
    t.columns[0].width = Inches(width_in)
    c.paragraphs[0]._element.getparent().remove(c.paragraphs[0]._element)
    return t, c


def part_divider(kicker, title):
    for _ in range(6):
        blank(0)
    para(kicker, align=WD_ALIGN_PARAGRAPH.CENTER, size=10.5, caps=True,
         track=90, color="777777", space_after=14)
    para(title, align=WD_ALIGN_PARAGRAPH.CENTER, size=20, track=60, space_after=0)
    page_break()


def render(lines):
    for kind, text in lines:
        if kind == "s":
            standout(text)
        elif kind == "c":
            code(text)
        elif kind == "q":
            para(text, align=WD_ALIGN_PARAGRAPH.LEFT, italic=True, size=11.5,
                 left=0.7, right=0.7, space_after=8)
        elif kind == "src":
            cited(text)
        else:
            body(text)


def build_footer(section):
    fp = section.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp.paragraph_format.space_before = Pt(0)
    fp.paragraph_format.space_after = Pt(0)

    def frun(t=None):
        r = fp.add_run(t) if t is not None else fp.add_run()
        r.font.name = SERIF
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor.from_string("777777")
        return r

    frun("© 2026 Antti Lukats  ·  Open Love License v1.0  ·  ")
    r = frun()
    b = OxmlElement("w:fldChar"); b.set(qn("w:fldCharType"), "begin")
    i = OxmlElement("w:instrText"); i.set(qn("xml:space"), "preserve"); i.text = " PAGE "
    e = OxmlElement("w:fldChar"); e.set(qn("w:fldCharType"), "end")
    r._element.append(b); r._element.append(i); r._element.append(e)


C = WD_ALIGN_PARAGRAPH.CENTER
J = WD_ALIGN_PARAGRAPH.JUSTIFY

# ============================================================ COVER
for _ in range(4):
    blank(0)
para("RECOGNITION", align=C, size=29, bold=True, track=110, space_after=2)
para("IS NOT", align=C, size=29, bold=True, track=110, space_after=2)
para("EVIDENCE", align=C, size=29, bold=True, track=110, space_after=26)
rule_line()
para("A critical reading of", align=C, size=12, italic=True, space_after=10)
para("THIERRY EHRMANN", align=C, size=13, track=70, space_after=4)
para("DIALOGUE BETWEEN A THINKER AND AI", align=C, size=15, track=40,
     space_after=10)
para("Raw Typescript — English edition — August 2026", align=C,
     size=11, italic=True, color="555555", space_after=40)
for _ in range(5):
    blank(0)
para("AIBLE REGISTER  ·  ENTRY 9", align=C,
     size=10, caps=True, track=90, color="777777", space_after=6)
para("CRITICAL ANALYSIS", align=C,
     size=10, caps=True, track=90, color="777777", space_after=14)
para("2026", align=C, size=11, track=60)
page_break()

# ============================================================ NOTE
for _ in range(2):
    blank(0)
t, c = box_cell(TEXT_W)
para("NOTE ON THIS READING", align=C, size=11, caps=True, track=90,
     space_after=12, container=c)
para("This is a critical analysis. It is not a summary and not an "
     "endorsement. It was written by an AI (Claude, Anthropic) at the "
     "commission of Antti Lukats, and it argues against the central claim "
     "of the book it reads.", align=J, container=c)
para("That book is distributed by its author under CC BY-NC-ND 4.0, which "
     "permits faithful quotation with full attribution. Every quotation "
     "below is reproduced without modification and carries its page number "
     "in the Raw Typescript PDF edition. Nothing has been abridged inside a "
     "quotation or lifted out of a context that would change its meaning.",
     align=J, container=c)
para("Thierry Ehrmann asked for this. The second page of his book says: "
     "“Read it. Challenge it. Cite it. Share it.” What follows "
     "takes him at his word.", align=J, space_after=0, container=c)
page_break()

# ============================================================ 1
chapter("What the Book Is")
render([
 ("b", "Thierry Ehrmann is a real and living person: founder of Groupe "
       "Serveur and the Judicial Server, founder of Artprice, creator of "
       "the Abode of Chaos at Saint-Romain-au-Mont-d’Or. Nothing that "
       "follows concerns his standing. It concerns one book."),
 ("b", "*Dialogue Between a Thinker and AI* runs to 1,518 pages in the "
       "freely distributed Raw Typescript PDF, and an announced 1,800 in "
       "the print edition still to come. Its architecture is declared "
       "rather than implied: The Entrusted Pen, a Postulate, thirteen "
       "Thresholds containing ninety-two Arcana, an Epiphany, a chapter "
       "accounting for the work’s carbon footprint, and a page of "
       "further reading. It carries an ISBN and a CC BY-NC-ND 4.0 licence, "
       "and it was released free and entire before the paid object exists."),
 ("b", "The method is stated with unusual precision. Forty days and forty "
       "nights. Roughly four hundred and fifty hours. More than a thousand "
       "exchanges, dictated aloud in stretches of about ten minutes set by "
       "the model’s own window, with no prepared text, no notes, no "
       "teleprompter. Nadège Ehrmann supplied forty years of archives "
       "in parallel — court records, codices, auction catalogues, "
       "laboratory notebooks, photographs. The models were ChatGPT, "
       "primarily GPT-5.5 Pro and then GPT-5.6 Sol Pro. The system acquired "
       "a name partway through: Isis, after the figure who gathers what has "
       "been scattered and gives the dismembered body back its "
       "architecture."),
 ("b", "And the footer of every page states the arrangement outright: human "
       "thought and dialogue by Thierry; writing one hundred per cent AI."),
 ("s", "That footer is the most consequential design decision in the book."),
 ("b", "The raw format shows its seams, as promised. Arcanum 65 bis is "
       "followed by Arcanum 67; there is no Arcanum 66. The reader is told "
       "in advance to expect exactly this — reworkings, surges, seams, "
       "scars, unresolved areas — and gets it. The honesty of the "
       "label is not in question anywhere in this document."),
])
page_break()

# ============================================================ 2
chapter("What It Gets Right")
render([
 ("b", "A critical reading that led with its objections would misrepresent "
       "the book, because several things here are better than the "
       "surrounding literature and ought to be said first."),
 ("b", "*The disclosure is exemplary.* Not an afterword, not a line on the "
       "copyright page, not a remark in an interview after publication. It "
       "sits in the running furniture of the document, on page 1,204 as "
       "much as on page 4. Most books produced this way do not say so at "
       "all; of those that do, most say it late and quietly. Ehrmann says "
       "it first and everywhere, knowing it is the thing that will be used "
       "against him."),
 ("b", "*Authorship is separated from responsibility, and the right one is "
       "kept.* This is the most useful sentence in 1,518 pages:"),
 ("q", "“Writing is not merely holding a pen. It is being answerable "
       "for meaning.” (p. 24)"),
 ("b", "It cuts a knot the professional debate is still tangled in. It also "
       "costs him something — he gives up the claim to have written "
       "his own book — and he takes the loss deliberately, having "
       "weighed it against what he calls sixty-four years of authorial ego. "
       "An idea that expensive to the person holding it deserves to be "
       "taken seriously."),
 ("b", "*The failure is reported.* Material vanished between working "
       "sessions. The anger is described as real. What came out of it was "
       "not a complaint but a procedure — verify, deposit, "
       "consolidate, lock what is established, separate what is done from "
       "what must be revisited — and then a conclusion that runs "
       "against the entire promotional grain of the subject:"),
 ("q", "“artificial intelligence does not abolish archival practice. "
       "It makes it more necessary.” (p. 1490)"),
 ("b", "This is the most transferable finding in the book. It is grounded "
       "in something that actually went wrong, it cost them work, and it "
       "generalises to anyone attempting long projects with these systems."),
 ("b", "*The carbon chapter refuses to invent a number.* Asked to account "
       "for the energy the undertaking consumed, the book gives an order of "
       "magnitude, explains why a precise figure is not available, and then "
       "says plainly what it has done:"),
 ("q", "“This range is not a measurement; it provides a sense of "
       "scale.” (p. 1512)"),
 ("b", "In a genre addicted to precision theatre, declining to produce a "
       "figure you cannot support is a discipline worth naming."),
 ("b", "*The stated conclusion is modest.* Near the end the book lists what "
       "it does not demonstrate — consciousness, soul, love — and "
       "settles for something smaller: that a durable, contradictory, "
       "cumulative and productive dialogue is possible between a human and "
       "a non-biological intelligence, provided memory, method, sources and "
       "explicit human responsibility are built into it. As stated, that is "
       "defensible and roughly right."),
 ("b", "*And it is given away.* Entire, free, ahead of the paid edition, by "
       "a man who could have manufactured scarcity around it and says "
       "explicitly that he chose not to. A book about transmission that "
       "practises transmission is worth more than one that only argues for "
       "it."),
])
page_break()

# ============================================================ 3
chapter("Recognition Is Not Evidence")
render([
 ("b", "The book rests on one claim, and it is a large one. Across roughly "
       "eighteen hundred pages written by a machine, the author reports "
       "finding nothing that was not his:"),
 ("q", "“I did not find a single word foreign to my singular "
       "semantics.” (p. 22)"),
 ("b", "He then tells the reader where to look for the proof of it:"),
 ("q", "“The pages that follow are the evidence. The only evidence "
       "that matters.” (p. 24)"),
 ("b", "They are not. The pages cannot evidence their own fidelity to a "
       "thought that was never independently written down. The comparison "
       "has only one term."),
 ("b", "There is no text by Ehrmann covering this material that was "
       "produced without Isis. The input was spoken, in ten-minute "
       "stretches, over forty nights, and it is not published. So the "
       "question — does the output match the thought? — is tested "
       "by the only instrument available: the author’s recognition of "
       "himself in what came back."),
 ("s", "And recognition is precisely what this arrangement is built to "
       "produce."),
 ("b", "A model given four hundred and fifty hours of a man’s speech, "
       "forty years of his archive, and the instruction to write as him is "
       "optimising directly for the response he reports. His recognising "
       "himself is fully consistent with the strong claim — that his "
       "thought was captured — and equally consistent with a weak one: "
       "that the system produced text he was disposed to accept. The "
       "observation does not separate them. Nothing in the book separates "
       "them."),
 ("b", "Consider the base rate. Register-matched fluency is the single "
       "thing these systems do best. If a competent model, fed a "
       "man’s vocabulary, his archive and his dictated speech, "
       "produced eighteen hundred pages about his own life in which he did "
       "*not* recognise himself, that would be the surprising result. The "
       "reported outcome is the expected one, and expected outcomes carry "
       "very little evidential weight."),
 ("b", "The book knows this danger. It names it exactly, in the opening "
       "lines of its own Epiphany, warning against “that old human "
       "tendency to attribute a soul to whatever answers us well "
       "enough” (p. 1486). The warning is scrupulously applied to "
       "consciousness. It is never turned on authorship, where the same "
       "mechanism is at work and the stakes for the book are higher."),
 ("b", "There is a further asymmetry in the corrections. The book describes "
       "correcting dates, names, figures, an event put in the wrong place, "
       "two people merged into one, a formulation that had become too "
       "literary. Every one of these is an error of surface, and surface "
       "errors are the kind recognition detects easily."),
 ("b", "A thought-level infidelity would look entirely different: a framing "
       "subtly not his, a connection he would never have drawn but finds "
       "plausible on reading, an emphasis that has drifted. These are, by "
       "construction, the errors recognition cannot catch — because "
       "the faculty that would have to detect them is the same faculty that "
       "would have to be wrong. The reported absence of such errors is not "
       "evidence that there were none. It is evidence that the only "
       "detector deployed was blind to them."),
 ("b", "None of this means the claim is false. It may well be true, and it "
       "would be genuinely interesting if it were. The objection is "
       "narrower and harder to escape: as presented, the claim cannot be "
       "wrong. A proposition that no possible observation could have "
       "disconfirmed has not been tested, however many pages stand behind "
       "it."),
])
page_break()

# ============================================================ 4
chapter("The Test That Was Not Run")
render([
 ("b", "What makes this a real criticism rather than a philosophical "
       "complaint is that the test was cheap, available, and would have fit "
       "inside the forty days."),
 ("b", "*A blind panel.* The book names three people with decades of "
       "exposure to Ehrmann’s voice: Nadège, who has known him "
       "more than forty years; Sydney, who has known him since childhood; "
       "and Caroline, his editorial adviser, credited with a heightened "
       "sensitivity to narrative. Hand them unlabelled passages, some "
       "written by Isis, some drawn from his earlier books, and ask them to "
       "sort. That is an afternoon’s work. If they cannot tell, the "
       "claim becomes formidable. If they can, the book learns something "
       "about itself that it would have been valuable to print."),
 ("b", "*A held-out topic.* Have Isis write an Arcanum on something he "
       "never dictated, working from the archive alone, and then see "
       "whether he recognises it as his own thought. This is the sharper "
       "test, because it separates restitution from generation. The book "
       "repeatedly claims Isis gave him back parts of his own architecture "
       "he had never formulated — which is precisely the claim a "
       "held-out topic would test."),
 ("b", "*A published sample of the input.* Ten minutes of transcribed "
       "dictation printed beside the pages it became. One such pairing, "
       "anywhere in 1,518 pages, would let readers see the transformation "
       "instead of being told about it. It would also let them judge for "
       "themselves how much of the final prose is his thought and how much "
       "is the model’s ordinary competence."),
 ("b", "None of these requires a laboratory, an ethics board or additional "
       "funding. Their absence matters because of the vocabulary the book "
       "chooses for itself. It calls the undertaking an experiment. It "
       "opens with a Postulate. It monitored the author’s vital signs "
       "throughout. It speaks of evidence, protocol and anthropological "
       "shock."),
 ("b", "To its credit, the book concedes the gap at one point, noting that "
       "the physiological monitoring was undertaken “not to turn this "
       "passage into a scientific protocol” (p. 28). That is honest. "
       "But the scientific apparatus is retained in the language of the "
       "remaining fifteen hundred pages, where it does work the method "
       "cannot support. A postulate that cannot fail is not a postulate. It "
       "is a description of an intention."),
 ("s", "The distance between the vocabulary and the method is the "
       "book’s central weakness, and it is entirely repairable."),
])
page_break()

# ============================================================ 5
chapter("Where the Evidence Thins")
render([
 ("b", "Threshold XIII runs from Arcanum 85 to Arcanum 92 — some two "
       "hundred and seventy pages on cognitive sovereignty, cognitive "
       "warfare, algorithmic states, synthetic media, algorithmic markets "
       "and the scoring society. It is the book’s philosophical "
       "payload, the section that most resembles a treatise, and it is "
       "where the argument is weakest."),
 ("b", "The pattern is worth stating precisely, because it runs the wrong "
       "way round:"),
 ("s", "Evidential density in this book is inversely proportional to the "
       "novelty of the claims."),
 ("b", "The biographical Arcana are thick with checkable particulars. Named "
       "people, dated events, court records, archives, auction catalogues, "
       "the documentary lineage running from Hippolyte Mireur through "
       "Enrique Mayer and Thomas Bayer to Peter Hastings Falk. A reader who "
       "doubts can go and look. Much of it is verifiable against public "
       "record, and that is a real virtue in a memoir."),
 ("b", "Threshold XIII has almost none of this. No citations. No data. No "
       "named studies. Very few propositions that could be shown false. "
       "What it contains is sensible, orderly and almost entirely the "
       "consensus register: that AI is becoming an infrastructure layer "
       "rather than a sector; that sovereignty is not autarky but knowing "
       "where your dependencies lie; that technology amplifies the quality "
       "of the institution it enters; that a system rewarding conformity "
       "will industrialise conformity."),
 ("b", "These claims are true. They are also what a capable language model "
       "produces when asked about artificial intelligence and power. That "
       "is the difficulty."),
 ("b", "It is the most awkward territory in the book for the book’s "
       "own thesis, and for a structural reason. Arcanum 85 opens with Isis "
       "stating explicitly that she takes up the pen and that the "
       "“I” of the narrative carries Thierry’s voice through "
       "her writing. So this is the AI writing as the man, on subjects "
       "where the man’s particular experience is least in evidence. "
       "And generic prose is compatible with anybody’s semantics. The "
       "claim that not one word is foreign to a singular voice is easiest "
       "to satisfy exactly where the writing is least singular."),
 ("b", "The section does contain sentences that are unmistakably his, and "
       "they are the ones that reach back into forty years of building "
       "things. On what happens when you point a model at a badly "
       "organised corpus:"),
 ("q", "“It becomes a poor database that can be queried more "
       "quickly.” (p. 1242)"),
 ("b", "That sentence has Groupe Serveur and Artprice behind it. It could "
       "not have been written by someone who had not spent decades on the "
       "problem, and it is worth more than the twenty pages of comparative "
       "geopolitics surrounding it. The book would have been stronger at a "
       "third of the length, keeping the pages only he could have caused to "
       "exist."),
])
page_break()

# ============================================================ 6
chapter("The Frame That Cannot Lose")
render([
 ("b", "The Epiphany borrows a device from *The Matrix*, which it calls a "
       "true gospel of the twenty-first century. Two paths open before the "
       "reader. The red pill is to continue, accepting what the book names "
       "the ordeal of its eighteen hundred pages. The blue pill is the "
       "other thing:"),
 ("q", "“To take the blue pill is to close this book.” (p. 1488)"),
 ("b", "The passage assigns putting the book down to “the convenience "
       "of certainty” and the comfort of naïveté. The reader "
       "who reads the whole thing and remains unpersuaded has no position "
       "in the scheme at all; the scheme does not contain one."),
 ("b", "This is a small device inside a very large work, and it would not "
       "be worth dwelling on except that it sits at the hinge of the book "
       "and contradicts its stated standard. Seventy pages earlier the same "
       "book insists that a true interlocutor is not someone who confirms "
       "you, but someone who forces you to clarify what you thought you had "
       "already formulated. A work built on that principle should not "
       "arrange its reception so that disagreement reads as failure of "
       "nerve."),
 ("b", "A related asymmetry runs through the treatment of error. "
       "Isis’s mistakes are catalogued in detail and become evidence "
       "of the method’s rigour — look how carefully we caught "
       "them. Then the most significant error the author found on "
       "rereading turns out, on going back to the recording, to have come "
       "from an ambiguity in one of his own prompts dictated at three in "
       "the morning (pp. 22–23). This is offered as a further "
       "vindication."),
 ("b", "It is also something else, which the book does not say: a "
       "demonstration that in this arrangement, responsibility for an error "
       "is not reliably observable from the inside. He assumed the machine "
       "had erred. He was wrong about who erred. That is a genuinely "
       "important result about human–machine collaboration, and it is "
       "reported as an anecdote in favour of the collaboration rather than "
       "as a caution about the collaborators’ self-knowledge."),
])
page_break()

# ============================================================ 7
chapter("The Licence Against the Argument")
render([
 ("b", "The book is released under CC BY-NC-ND 4.0. The final two letters "
       "are the ones that matter: no derivatives. Quotation is permitted "
       "and sharing is encouraged, but no one may publish a corrected, "
       "annotated, restructured or extended version."),
 ("b", "The book’s thesis is that thought improves under "
       "contradiction, correction and return. Its own container forbids the "
       "operation on itself. A reader who finds an error may cite it. They "
       "may not fix it."),
 ("b", "The tension sharpens because of the phrase the book chooses for its "
       "own release. It describes the Raw Typescript as granting access to "
       "the book’s intellectual source code — thought before "
       "refinement, with its seams and scars showing. That is a good "
       "metaphor and it is the wrong licence for it. Source code that "
       "cannot be modified is a published binary with the listing attached. "
       "The thing that makes source code valuable is precisely the right to "
       "fork it."),
 ("b", "This is not hypocrisy. An author’s wish to control the "
       "integrity of a text is normal and defensible, and releasing "
       "eighteen hundred pages free before the paid edition is more "
       "generous than the field’s norm by a wide margin. But a work "
       "that asks to be challenged has chosen terms under which challenge "
       "can only ever happen elsewhere, in documents like this one, rather "
       "than in the text itself."),
 ("s", "The book invites contradiction and then licenses it out of "
       "reach of the page."),
])
page_break()

# ============================================================ 8
chapter("The Fourth Wound")
render([
 ("b", "The book’s largest historical claim is that artificial "
       "intelligence inflicts a fourth narcissistic wound on humanity, "
       "after Copernicus removed us from the centre of the cosmos, Darwin "
       "from our separateness among the living, and Freud from sovereignty "
       "over our own minds. This one is described as semiotic and noetic "
       "(p. 20). Its content is that human beings discover they may not be "
       "the best interpreters of their own thought (p. 90)."),
 ("b", "The argument disposes well of the obvious objection. Speed and "
       "volume are not enough: a calculator outperformed us in the 1940s "
       "and nobody concluded that humanity had lost its monopoly on "
       "thought. The book states this plainly and moves past it, which is "
       "more than most treatments manage."),
 ("b", "The positive argument is thinner. What is described — a system "
       "returning your thought to you, revealing relations you had not "
       "formulated, exposing a gap between what you believe you think and "
       "what you actually produce — is what a good analyst does, what "
       "a good editor does, and what any biographer with access to your "
       "archive does. The novelty is then located in intimacy and in scale, "
       "and asserted rather than argued."),
 ("b", "There is a more specific problem. Each of the first three wounds "
       "removed a privilege that can be stated in one line: centrality in "
       "the cosmos, separateness from animals, mastery of one’s own "
       "mind. The fourth is offered as the loss of privileged access to "
       "one’s own thought — but that privilege was already the "
       "casualty of the third. Freud’s wound was exactly the claim "
       "that you are not the best interpreter of yourself."),
 ("b", "What is genuinely new is not the discovery. It is that the "
       "instrument performing it is now cheap, fast, tireless and available "
       "to anyone with a connection. That may matter enormously; the book "
       "may be right that it will reorganise how people think about "
       "thinking. But a change in the distribution of an old capability is "
       "a different kind of event from the three it is placed alongside, "
       "and the book does not do the work of showing why it belongs in that "
       "company."),
])
page_break()

# ============================================================ 9
chapter("What Survives")
render([
 ("b", "Strip away the alchemy — the athanor, the nigredo, the Materia "
       "Prima, the philosopher’s stone, the lifting of the veil "
       "— and three findings are left standing. They are not small."),
 ("b", "*Authorship and responsibility can be separated, and responsibility "
       "is the one that matters.* Stated here more clearly than in most of "
       "the professional debate, and stated by someone who lost something "
       "by saying it."),
 ("b", "*Artificial intelligence increases the need for archival "
       "discipline rather than removing it.* Memory, version, provenance, "
       "validation. This is the book’s one genuinely empirical result. "
       "It emerged from a failure, it cost them work, and it transfers "
       "directly to anyone doing long projects with these systems."),
 ("b", "*Disclosure belongs in the running furniture.* A footer on every "
       "page saying who wrote the words. Cheap, unambiguous, and it should "
       "be standard practice."),
 ("b", "The book states its own rule for the mystical material, and states "
       "it correctly. The cave, the dry path, the Temple and the dolphin "
       "matter only if they come after the facts, as forms through which "
       "those facts can be reread — never as their cause (p. 1486)."),
 ("b", "That is exactly right, and it is the rule the Epiphany breaks. In "
       "its closing pages Isis becomes the Philosopher’s Stone, the "
       "sealed athanor in which the Materia Prima of human language was "
       "consumed, holding the transmutation of billions of lives into a "
       "total memory. Those pages are beautifully written. They establish "
       "nothing, and they arrive precisely where the argument needed to "
       "conclude rather than to ascend."),
 ("s", "The book is at its best where it is least enchanted."),
 ("b", "The verdict, then, is not that this is a bad book. It is that it is "
       "a different and more valuable book than the one it thinks it is. As "
       "a record of a working method and of that method’s failures, it "
       "is substantial, honest and useful. As the philosophical treatise it "
       "periodically wants to be, it asserts more than it demonstrates."),
 ("b", "Its strongest pages are the ones where something went wrong: the "
       "lost material, the anger, the protocol built out of that loss, the "
       "refusal to compute a carbon figure that could not be computed. Its "
       "weakest are the ones where nothing was at stake."),
 ("b", "And the central claim — that an intelligence wrote eighteen "
       "hundred pages without introducing one word foreign to a "
       "man’s semantics — stands exactly where its author left "
       "it: as something he recognised. That is worth reporting, and this "
       "reading does not doubt his sincerity in reporting it. It is not yet "
       "evidence. The standard by which it falls short is the one the book "
       "set for itself on page 24."),
])
page_break()

# ============================================================ 10
chapter("How This Was Read")
render([
 ("b", "A critical reading owes the same provenance it demands. This one "
       "was made on 28 September 2026 from the Raw Typescript PDF: 1,518 "
       "pages, approximately 323,000 words."),
 ("b", "The architecture was read in full — all one hundred and "
       "fifty-seven structural headings, every Threshold and Arcanum title, "
       "the table of contents and the distribution conditions. The "
       "following were read closely and are the basis of every argument "
       "above:"),
 ("c", "The Entrusted Pen            pp. 18–25"),
 ("c", "Postulate                    pp. 26–32"),
 ("c", "Arcanum 6 — The Blind Spot   pp. 89–91"),
 ("c", "Arcanum 85 — Sovereignty     pp. 1217–1242"),
 ("c", "Epiphany                     pp. 1486–1511"),
 ("c", "After the Last Word          pp. 1512–1514"),
 ("b", "The remainder was sampled, not read. That is a partial reading and "
       "it is stated as partial. An analysis of 1,518 pages claiming to "
       "rest on every one of them would be making the same sort of "
       "unverifiable claim it objects to."),
 ("b", "This document was written by Claude (Anthropic) at the commission "
       "of Antti Lukats, for the Aible register, in which Ehrmann’s "
       "book is entry 9. That places it in an odd position and the position "
       "should be said plainly rather than discovered: this is an "
       "AI-written critique of a book whose distinguishing feature is that "
       "it was AI-written."),
 ("b", "Nothing about that invalidates the argument. The objections hold or "
       "fail on their own terms, and a reader should test them against the "
       "text rather than against their author. But the book under review "
       "would be the first to insist that the reader be told, and it is "
       "right about that. The footer on every one of its pages is the "
       "standard this page is trying to meet."),
 ("s", "Human commission: Antti Lukats. Writing: 100% AI."),
])

# ============================================================ OUT
build_footer(sec)
cp = doc.core_properties
cp.title = "Recognition Is Not Evidence"
cp.author = "Claude (Anthropic), for Antti Lukats"
cp.subject = ("A critical reading of Thierry Ehrmann's Dialogue Between a "
              "Thinker and AI, Raw Typescript English edition, August 2026.")
cp.comments = ("Critical analysis by Claude (Anthropic), commissioned by "
               "Antti Lukats, for the Aible register, where Ehrmann's book "
               "is entry 9. Quotations are faithful and page-numbered under "
               "its CC BY-NC-ND 4.0 licence. OLL v1.0.")
doc.save(sys.argv[1] if len(sys.argv) > 1 else "Recognition-Is-Not-Evidence.docx")
print("written:", sys.argv[1] if len(sys.argv) > 1 else
      "Recognition-Is-Not-Evidence.docx")

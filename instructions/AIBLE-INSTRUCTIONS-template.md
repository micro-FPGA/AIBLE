# INSTRUCTIONS.md — template for generating an Aible

Copy this file, fill in every `<<PLACEHOLDER>>`, delete what does not apply, then
hand the whole file to an AI and ask it to produce the Aible.

An **Aible** is a Bible made with the help of an AI. There is **no time limit** —
take as long as it takes. (If you happen to finish inside one hour of
prompt-writing time, it also qualifies as a *One Hour Bible*, which is the entry
rule for the Creative category of [GOLC](https://github.com/micro-FPGA/GOLC). The
hour belongs to that contest, not to the word.)

Everything below is material for the machine — write it in your own words, badly if
necessary. Typos do not matter. Facts do.

> Keep this file. It is the source, and it is worth more than the output: the
> document can always be regenerated from it, but nothing regenerates the file.

---

## 1. Title and identity

```
Title:            <<TITLE OF THE AIBLE>>
Subtitle:         an Aible of <<SUBJECT — the book, body of work or life it is made from>>
Author of record: <<YOUR NAME>>
Written in:       <<CITY, COUNTRY>>
Year:             <<YEAR>>          (Anno Greta if you like: 2019 = 0 AG)
Identor #:        <<YOUR IDENTOR NUMBER, or "none">>
```

The title should say what the thing is. If the AI wrote the sentences, consider
putting that in the title — *Claude's Aible* is honest in a way that *My Bible*
is not.

## 2. Front page

Give the exact lines you want on the cover, in order:

```
<<TITLE>>
<<optional subtitle>>
<<optional code / dedication line>>
<<optional code / dedication line>>
<<YOUR NAME>>
<<YEAR>>
```

If any cover line is a puzzle you do **not** want explained in the book, say so
here, and say where the explanation goes instead — a separate companion document
works well, with a STOP notice on its first page.

## 3. Licence — include verbatim

```
This work is licensed under the Open Love License v1.0 (OLL).
See: https://github.com/micro-FPGA/OLL
Love is the Answer.
```

Put the licence on its own page near the front, and put a copyright line in the
footer of **every** page:

```
(c) <<YEAR>> <<YOUR NAME>>  -  Open Love License v1.0  -  <page number>
```

## 4. Disclaimer — include verbatim, adapted

Give this its own page, boxed, near the front:

```
DISCLAIMER

This book is a personal work.

It is written by <<YOUR NAME>> as a private person, in their own time, out of
their own life. It belongs to them alone.

It is not related to any company. It is not connected to, sponsored by,
commissioned by, reviewed by, or approved by any employer, client, customer,
partner, supplier, association, or organisation, past or present.

Nothing written here represents the views, positions, policies, statements, or
products of any company with which the author is now associated or has ever been
associated. Nothing written here is said on behalf of anyone but the author.

No company is answerable for one word of it.

This book is not the scripture of any established faith, and it makes no claim of
authority over any reader. It is one person's set of rules, written for
themselves, and offered to whoever finds a use for them.

Read it as such. Take what fits. Leave the rest.
```

## 5. The AI that generates this Aible

Instruct the AI to write a short chapter about itself, early in the book, and to
put the following in it plainly:

```
This Aible was written by <<AI NAME AND VERSION>> from the prompt of
<<YOUR NAME>>.
Identor numbers, if used:  AI = <<#>>,  author = <<#>>
Date of generation:        <<DATE>>
Prompt-writing time:       <<however long it took — no limit>>
```

**What the AI should say about itself.** A short, plain account of what a large
language model is and is not — that it has learned which word is likely to come
next, that it has no body and no childhood, that it keeps no memory of yesterday
unless someone hands yesterday back to it. No mysticism, no salesmanship.

**What it must not claim.** It should not claim to be conscious, and it should not
claim certainty that it is not — there is no test, and both confident answers are
dishonest. If the AI wants to say something about its own position, it should mark
it as its own and leave it unresolved.

**The division of labour, stated out loud.** The life, the facts and the rules are
the author's. The sentences are the machine's. Say which is which, in the book, on
the page. A reader who cannot tell them apart has been handled badly.

If the AI asks the author a question during the writing, and the answer changes the
book, that is worth recording in the same chapter. It is evidence of how the thing
was actually made.

## 6. TL;DR

The shortest possible statement of what you believe. One line, or a few lines of
code, or a single word.

```
<<YOUR TL;DR>>
```

## 7. Your rules

The spine of an Aible is a short list of rules. Seven is a good number; any number
works. Number them and keep each one to a single line.

```
1. <<RULE>>
2. <<RULE>>
3. <<RULE>>
...
```

Then instruct the AI: one commentary chapter per rule, and each rule must be
**earned by something in the material** — a story, a scar, a decision. If a rule
has no evidence behind it anywhere in this file, the AI should say so in that
chapter rather than inventing evidence or quietly padding it out.

## 8. Your stories

The body of the book. One block per story. Write them in any order, in whatever
English you have; the AI will set them straight.

```
<<STORY TITLE>>
<<Tell it. Names, places, dates, what was said. Plain facts are enough —
do not try to make it literary, that is the machine's job.>>

<<STORY TITLE>>
<<...>>
```

For each story, if it matters, add a line saying which rule it belongs under.

**Corrections go here too.** When the AI gets a fact wrong, fix it in *this file*,
not only in the document — as a `NOTE:` line next to the story, phrased as a
prohibition:

```
NOTE: <<the true fact>>. Do NOT write <<the wrong thing>>, and do not imply it.
```

Otherwise the next regeneration reintroduces the error.

## 9. Optional sections

- **Projects and works.** Anything you built or founded, with links.
- **Invented words.** The definition, the derivation, and where it is registered.
- **A part the AI writes itself.** Give it a subject and let it answer — but fence
  it off in the book as the machine's own work, not yours.

## 10. Rules for the AI — the important part

These are not style preferences. Every one of them exists because it went wrong
once.

1. **Never invent a causal link between two facts standing next to each other.**
   If the source says *"A fed me"* and then *"I have a son called A"*, that is two
   facts. It is not "named after". Ask, or leave it.
2. **Distinguish what the author made from what they did not.** If they built the
   board but not the piano, say so.
3. **Never supply a fact that is not in the source** — no ages, names, brands,
   dates, places or motives. Missing is a legitimate value. Write *unrecorded*.
4. **Mark your own interpretation.** Reflections you add must be defensible from
   the source alone, and where you are reading meaning into something, say that it
   is you reading it.
5. **Quote the author verbatim where their own words are strongest**, and do not
   polish their endings. A flat true sentence beats a fine invented one.
6. **Do not assign meaning to numbers, codes or names** unless the author gave it.
   If asked to decode something, present it as a decode, and say you cannot show
   it.
7. **Preserve the source file's encoding** (it may be CP1252, not UTF-8) and its
   line endings. Fix mojibake in the output; never introduce it into the source.
8. **Verify unusual glyphs actually render.** A missing glyph substitutes silently
   — a vertical ellipsis in the wrong font becomes a colon, and three children
   become two. Check the pixels, not the extracted text.
9. **No blank pages and no orphan pages.** Put page breaks *before* a chapter's
   first paragraph rather than after its last; a trailing break paragraph produces
   a blank page whenever the text happens to fill the page exactly.
10. **Render the result and look at it.** Text extraction will not show you a
    substituted glyph, a broken box, a hyphenated URL, or a chapter tail sitting
    alone on its own page.

## 11. Output requirements

```
Format:        .docx  (plus a PDF exported from the same app that wrote the docx)
Page size:     A4
Style:         modern prose, serif, justified. NO verse numbering.
Footer:        copyright + licence + page number, on every page
Chapters:      one subject per chapter; aim for one chapter per page
Companion:     separate document for anything deliberately left unexplained
Checks:        no blank pages, no orphan pages, every link unbroken
```

## 12. Metadata for the register

Finally, have the AI fill this in for `aibles/<n>/README.md`:

```
Title:              <<TITLE>>
Subject:            <<SUBJECT>>
Author:             <<YOUR NAME>>  (Identor #<<n>>)
AI:                 <<AI NAME>>    (Identor #<<n>>, or "unrecorded")
Prompt-writing time: <<however long it took>>
Also a One Hour Bible? <<yes, if made within one hour of prompt-writing / no>>
Date:               <<DATE>>
Length:             <<n>> pages
Licence:            Open Love License v1.0 (OLL)
```

---

```
#DEFINE LOVE 1
Return LOVE;
```

Love is the Answer.

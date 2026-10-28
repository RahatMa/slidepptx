/* The prompt people paste into Claude together with question images or PDFs. Keep CLAUDE_PROMPT.md in sync. */
(function (root) {
  root.QB_CLAUDE_PROMPT = `I'm attaching physics questions (photos, screenshots or PDF pages). Convert them into a spreadsheet for my slide generator.

ExamID for every row: [WRITE YOUR EXAM ID HERE, e.g. PHY-01]
Solutions: [write "copy from source only" OR "write solutions where missing"]

OUTPUT
Create ONE .xlsx file with ONE sheet named "Questions". Row 1 must be exactly these 12 headers, in this order:
ExamID | QNo | Question | ImageURL | OptionA | OptionB | OptionC | OptionD | Answer | Solution | Type | Layout
One row per question, in the order they appear. No merged cells, no extra sheets, columns or notes.

TEXT
- Copy the Bengali text exactly as printed. Do not translate or rephrase. Fix only obvious OCR mistakes.
- QNo = the number printed in the source (digits only, e.g. 7).
- A new line inside a cell = a new line on the slide. Use it for statement lists (i., ii., iii.) and for solution steps.

MATH (important)
- Put every formula, variable, number-with-unit and symbol inside \\( ... \\). Never use $ ... $.
- Use LaTeX: powers ^{-2}, subscripts F_{\\mathrm{net}}, fractions \\frac{mv^2}{r}, roots \\sqrt{2gh}, degrees 60^\\circ, vectors \\vec{F}, \\times, \\Delta, \\theta, \\mu.
- Units upright with a space: \\(9.8\\ \\mathrm{m\\,s^{-2}}\\), \\(20\\ \\mathrm{N}\\).
- Never use Unicode superscripts (² ⁻¹) — always ^{ }.
- A main equation in a solution goes on its own line as \\[ ... \\].

MCQ
- Type = MCQ. OptionA–OptionD hold the option text WITHOUT the (ক)/(a)/A. prefix.
- Answer = one of ক খ গ ঘ (or A B C D). Use the key printed in the source. If there is no key, solve it yourself and tell me in your reply which answers you inferred.
- Statement MCQs: put the intro line, each statement (i., ii., iii.) and the final line "নিচের কোনটি সঠিক?" in the Question cell, each on its own line.

WRITTEN / CREATIVE PROBLEMS
- Type = Problem. Leave OptionA–OptionD empty. Answer = the final result (optional), e.g. \\(\\mu = 2\\sqrt{3}\\).

FIGURES
- If a question has a diagram, crop it from the page into a PNG (tight crop, white background, at least 800 px wide, no question text inside it).
- Name it <ExamID>_<QNo>.png (e.g. PHY-01_24.png) and write that exact file name in ImageURL.
- Give me all PNGs together in ONE .zip file. If you cannot crop a figure, write NEED_IMAGE in ImageURL.

LAYOUT column: leave empty.

When done, give me the .xlsx and the .zip, then a short list of: number of questions, anything unreadable, and answers you inferred.`;
})(typeof self !== 'undefined' ? self : this);

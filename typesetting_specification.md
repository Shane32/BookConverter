# Typesetting Specification for Only an Irish Boy

A complete style and layout reference for Microsoft Word typesetting of the novel Only an Irish Boy on 6" × 9" trim size.

---

## 📄 Page Layout

### Page Size and Margins
- **Trim Size**: 6" × 9"
- **Top Margin**: 1.0"
- **Bottom Margin**: 0.75"
- **Inside (Gutter)**: 1.0"
- **Outside**: 0.75"
- **Header Distance from Top**: 0.5"
- **Footer**: Not used (no content in footer)

### Section Formatting
- **Chapters**: Each chapter starts on a new recto (right-hand) page
- **Dedication**: Follows the title page on a recto (right-hand) page
- **Blank Pages**: Configurable via `FORCE_BLANK_VERSO_PAGES` flag in converter:
  - When `True`: Force blank verso (left-hand/even) pages before each chapter begins on an odd page; if the prior chapter ends on an even page, there will be two blank pages before the new chapter
  - When `False` (default): Allow chapters to end on an even page immediately preceding a new chapter (without forcing blank pages between them)
  - Note: Chapters will still always start on odd (right-hand) pages regardless of this setting
- **Section Breaks**: Use section breaks between chapters to change header text
- **Different Odd and Even Headers**: Enabled
- **Different First Page**: Enabled (title, dedication, and chapter first pages have no header)

---

## 🧾 Styles

### Title Page

#### `BookTitle` (custom)
- **Use**: Main book title
- **Font**: Georgia, 16 pt, Bold, All Caps
- **Alignment**: Centered
- **Vertical Alignment**: Centered on page

#### `BookSubtitle` (custom)
- **Use**: Author name with "by" prefix
- **Font**: Georgia, 12 pt
- **Alignment**: Centered
- **Spacing After**: 36 pt

---

### Dedication Page

#### `DedicationTo` (custom)
- **Use**: "To" line in dedication
- **Font**: Georgia, 12 pt, Bold
- **Alignment**: Centered
- **Spacing After**: 20 pt
- **Vertical Alignment**: Centered on page

#### `DedicationFrom` (custom)
- **Use**: "From" line in dedication
- **Font**: Georgia, 12 pt, Italic
- **Alignment**: Centered
- **Spacing After**: 20 pt

#### `DedicationCredits` (custom)
- **Use**: Credits in dedication
- **Font**: Georgia, 12 pt
- **Alignment**: Centered
- **Spacing After**: 20 pt

**Note**: A blank paragraph with DedicationFrom style is inserted before credits to provide additional spacing.

---

### Table of Contents

#### `TOCHeading` (custom)
- **Use**: "TABLE OF CONTENTS" heading
- **Font**: Georgia, 14 pt, Bold, All Caps
- **Alignment**: Centered
- **Spacing Before**: 36 pt
- **Spacing After**: 42 pt

#### `TOC Entry` (custom)
- **Use**: Chapter listings
- **Font**: Book Antiqua, 10 pt
- **Line Spacing**: 1.15
- **Spacing After**: 6 pt
- **Paragraph Indent**: 0.25"
- **Tabs**:
  - Left-aligned tab at **0.81"**
  - Right-aligned tab at **4.0"** with **dot leader**
- **Example Format**:
  ```
  I.<TAB>The New Home<TAB>1
  II.<TAB>A Friendly Call<TAB>7
  ```
- **Note**: The Table of Contents should end with a blank page

---

### Chapter Headings

#### `ChapterNumber` (custom)
- **Use**: Chapter number (e.g. "CHAPTER V")
- **Font**: Georgia, 14 pt, Bold, All Caps
- **Alignment**: Centered
- **Spacing Before**: 36 pt
- **Spacing After**: 14 pt
- **Keep with Next**: Enabled

#### `ChapterTitle` (custom)
- **Use**: Chapter title (e.g. "The Widow’s Proposal")
- **Font**: Georgia, 12 pt, Title Case
- **Alignment**: Centered
- **Spacing After**: 42 pt
- **Keep with Next**: Enabled

---

### Body Text

#### `Normal` (built-in)
- **Use**: All paragraphs of main text
- **Font**: Book Antiqua, 10 pt
- **Line Spacing**: 1.15
- **Spacing After**: 11.5 pt
- **First Line Indent**: 0.25"
- **Alignment**: Justified
- **Hyphenation**: Enabled
- **Widow/Orphan Control**: Enabled

#### `Quote` (custom)
- **Use**: Quoted text, letters, or other content requiring double indentation
- **Font**: Book Antiqua, 10 pt, Italic
- **Line Spacing**: 1.15
- **Spacing After**: 11.5 pt
- **Left Indent**: 0.25"
- **Right Indent**: 0.25"
- **Alignment**: Justified
- **Hyphenation**: Enabled
- **Widow/Orphan Control**: Enabled

---

### Headers

#### `PageHeader` (custom)
- **Use**: Single-line header on all pages except title and chapter starts
- **Font**: Book Antiqua, 9 pt, Italic
- **Alignment**: 
  - Chapter name left-aligned on verso, right-aligned on recto
  - Page number positioned via tab
- **Tab Stops**:
  - Right-aligned tab stop at **4.25"**
- **Example Content**:
  ```
  Chapter Title<TAB>123
  ```
- **Header Settings**:
  - **Different Odd and Even**: Enabled
  - **Different First Page**: Enabled (blank for title/chapter pages)

---

## ✅ Additional Recommendations
- Use `Section Break (Next Page)` before each chapter to assign header text.
- Use Word's `Link to Previous` header option selectively to manage running headers by chapter.
- Enable **Show/Hide ¶** to manage spacing and tab consistency.
- Use **fields** (like `{ PAGE }`) for dynamic page numbers.

---
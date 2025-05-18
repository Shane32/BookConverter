# Book JSON Format Documentation

This document describes the JSON format designed to store the content of "Only an Irish Boy; Or, Andy Burke's Fortunes" by Horatio Alger, Jr.

## JSON Structure

```json
{
  "book": {
    "title": "Only an Irish Boy",
    "subtitle": "Or, Andy Burke's Fortunes",
    "author": "Horatio Alger, Jr."
  },
  "dedication": {
    "to": "John Doe",
    "from": "Jane Doe",
    "credits": [
      "Cover art by Jack Doe"
    ]
  },
  "chapters": [
    {
      "number": 1,
      "title": "Andy Burke",
      "paragraphs": [
        "John, saddle my horse, and bring him around to the door.",
        "The speaker was a boy of fifteen, handsomely dressed, and, to judge from his air and tone, a person of considerable consequence, in his own opinion, at least. The person addressed was employed in the stable of his father, Colonel Anthony Preston, and so inferior in social condition that Master Godfrey always addressed him in imperious tones.",
        "John looked up and answered, respectfully:",
        {
          "type": "quote",
          "content": "MR. STONE: Sir—My son Godfrey complains that you have punished him severely for a very trifling fault. I wish to say that I consider your course very unreasonable and tyrannical, and desire you to understand that my son is not to be treated in such a manner."
        },
        // Additional paragraphs...
      ]
    },
    {
      "number": 2,
      "title": "A Skirmish",
      "paragraphs": [
        "Andy Burke was not the boy to run away from an opponent of his own size and age. Neither did he propose to submit quietly to the thrashing which Godfrey designed to give him. He dropped his stick and bundle, and squared off scientifically at his aristocratic foe.",
        // Additional paragraphs...
      ]
    },
    // Additional chapters...
  ]
}
```

## Structure Explanation

1. **Book Metadata**
   - `title`: The main title of the book
   - `subtitle`: The subtitle of the book
   - `author`: The author's name

2. **Dedication** (optional)
   - `to`: The person to whom the book is dedicated (optional)
   - `from`: The person dedicating the book (optional)
   - `credits`: An array of strings for additional credits (optional)

3. **Chapters Array**
   - Each chapter is an object containing:
     - `number`: The chapter number (integer)
     - `title`: The chapter title (string)
     - `paragraphs`: An array of strings, each representing a paragraph of text

4. **Paragraphs**
   - Each paragraph can be either:
     - A plain text string (for regular paragraphs)
     - An object with the following properties:
       - `type`: The type of paragraph (e.g., "quote")
       - `content`: The text content of the paragraph
   - All HTML formatting is removed
   - Original paragraph breaks are preserved
   - Special paragraph types:
     - `quote`: Used for quoted text, letters, or other content requiring double indentation

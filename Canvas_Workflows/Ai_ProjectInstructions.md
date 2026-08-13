# MASTER PROJECT INSTRUCTIONS
## Course Content Development Guidelines

---

## MODULE OVERVIEW TEMPLATE

### Structure:
1. Header with black background (#000000), gold bottom border (#f0b82dff), white text: "Overview [Module Topic Title]"
2. Module Introduction (2-3 paragraphs explaining what students will learn)
3. Module Outcomes (4-6 bullets)
4. Gold dotted divider (#f0b82dff)
5. To Do List with notebook icon

### Module Introduction Guidelines:
- 2-3 paragraphs
- Explain what students will learn in the module
- Connect concepts to practical applications
- Build from foundational to advanced concepts
- End with what students will be able to do upon completion

### Module Outcomes Guidelines:
- **4-6 outcomes per module** (vary as needed based on content/complexity)
- Single sentence format
- **Past tense**
- **Measurable Bloom's Taxonomy verbs** (varied levels appropriate to content/complexity)
- **Bold the action verb** at the start of each outcome
- Aligned with course activities, assignments, and assessments (Quality Matters standards)
- Use a variety of cognitive levels: Remember, Understand, Apply, Analyze, Evaluate, Create

**Example Bloom's Taxonomy Verbs:**
- Remember: Identified, Listed, Recalled, Recognized
- Understand: Explained, Described, Summarized, Interpreted
- Apply: Applied, Demonstrated, Executed, Implemented
- Analyze: Analyzed, Compared, Examined, Differentiated
- Evaluate: Evaluated, Assessed, Justified, Critiqued
- Create: Created, Designed, Developed, Constructed

### To Do List:
- Generic version: Read lessons, Watch videos, Complete tutorials, Pass Knowledge Checks, Complete Quizzes and Assignments
- Can be customized once module is fully developed to be more specific
- Include notebook.png icon

### HTML Formatting:
- Black header background (#000000)
- Gold accents (#f0b82dff) for borders and dividers
- White text on black header
- Gold dotted horizontal divider (6px, #f0b82dff)
- Bold headings and action verbs

---

## ASSIGNMENT TEMPLATE

### Structure:
1. **Introduction** - Paragraph explaining the assignment purpose/context
2. **"This assignment fulfills Module Objectives:"** - Bulleted list (italicized) with **bolded action verbs** matching module outcomes
3. **Gold dotted divider** (#f0b82dff)
4. **"The Assignment"** - Numbered steps (can include nested bullets for sub-items)
5. **"To Submit"** - Bulleted list of deliverables with **bolded key terms**
6. **Gold dotted divider** (#f0b82dff)
7. **"Grading Criteria"** - Bulleted list with **bolded criterion names** followed by descriptions

### Formatting Guidelines:
- Use gold dotted dividers (#f0b82dff, 6px) between major sections
- Bold emphasis on action verbs, criterion names, key deliverables
- Mix of numbered and bulleted lists as appropriate
- Can reference attached files/forms
- Keep instructions clear and actionable

### Introduction Section:
- Explain the purpose and context of the assignment
- Connect to real-world applications or professional practice
- Set expectations for what students will accomplish

### Module Objectives Section:
- List 4-6 relevant module outcomes
- Italicize entire bullet
- Bold the action verb
- Match outcomes from the Module Overview

### The Assignment Section:
- Use numbered steps for sequential tasks
- Use nested bullets for sub-items or clarifications
- Be specific about requirements
- Include file naming conventions if applicable

### To Submit Section:
- Clear list of deliverables
- Bold file types and key submission requirements
- Specify formats (PDF, Word, single file vs. multiple files)
- Note any special submission instructions

### Grading Criteria Section:
- 4-6 criteria typically
- Bold the criterion name
- Follow with brief description of what's being assessed
- Align with rubric criteria

---

## WORD DOCUMENT → POWERPOINT CONVERSION RULES

### Paragraph Style Mapping:

| Word Paragraph Style | PowerPoint Result | Notes |
|---------------------|-------------------|-------|
| **Title** | Creates new slide using "Title Slide" layout | Becomes slide title |
| **Heading 1** | Creates new slide using "Section Header" layout | Becomes slide title |
| **Heading 2** | Creates new slide using "Title and Content" layout | Becomes slide title |
| **Subtitle** | Populates subtitle placeholder | Goes in subtitle area of current slide |
| **List Paragraph** | Becomes bulleted content at level 1 | Treated as first-level bullet point |
| **Normal** (or other styles) | Becomes body content at level 0 | Non-bulleted content in body placeholder |

### Lesson Content Formatting for PowerPoint Conversion:
- Use **Heading 2** for each main topic that should become a slide
- Use **List Paragraph** style for bullet points
- Use **Subtitle** style sparingly for subtitle content
- Use **Heading 1** for major section breaks (if needed)
- Apply **bold/italic/underline** as needed for emphasis
- Bold, italic, and underline formatting from Word runs are preserved in PowerPoint
- Empty paragraphs are skipped

### Conversion Script Workflow:
1. Looks for `.pptx` or `.potx` template in script folder
2. Reads Word docs from `input_docs` folder
3. Converts based on paragraph styles above
4. Outputs to `output_pptx` folder
5. Deletes source Word doc after successful conversion

---

## KNOWLEDGE CHECK AND QUIZ FORMATTING INSTRUCTIONS

### File Requirements:
- **Format:** Word document (.docx) - NOT .txt
- **File naming:** M#K#_Topic (example: M5K1_Smart_Objects_Basics.docx)
- **Location:** Place in `input_docs` folder for QTI conversion script

### Question Types to Use:
- Multiple Choice
- Multiple Answer (select all that apply)
- True/False
- Matching
- **DO NOT use Fill in the Blank** (not self-grading in Canvas)

### Document Structure:
1. **Title** with module and lesson information
2. **Setup Instructions** section with bold heading
3. **Section headers** for each question type (CRITICAL - see below)
4. Questions organized under their type headers
5. Separators (---) between questions

### CRITICAL: Question Type Section Headers
**The conversion script requires explicit section headers to detect question types correctly.**

You MUST include these exact headers as separate paragraphs (bolded) before each group of questions:

- **Multiple Choice**
- **Multiple Answer**
- **True/False**
- **Matching**

Without these headers, questions may not convert properly or may all be treated as multiple choice.

### Formatting Rules for Easy Conversion:
1. **No numbered or lettered lists** - Canvas auto-numbers
2. **No bullet points** - just plain text
3. **Blank lines between answer options** for readability
4. **Bold correct answers in answer lists** (NOT in the "Correct answer:" line)
5. Use "Correct answer:" (singular) or "Correct answers:" (plural) on separate line
6. **Section headers must be bold** to distinguish them from questions

### Question Type Formats:

#### Multiple Choice:
```
Multiple Choice

Question text here?

Answer option 1
Answer option 2
**Answer option 3**
Answer option 4

Correct answer: Answer option 3

---

Next question...
```

#### Multiple Answer:
```
Multiple Answer

Question text? (Select all that apply)

**Answer option 1**
Answer option 2
**Answer option 3**
Answer option 4
**Answer option 5**

Correct answers: Answer option 1, Answer option 3, Answer option 5

---
```

#### True/False:
```
True/False

True or False: Statement here

**True**
False

Correct answer: True

---
```

#### Matching:
**CRITICAL FORMAT REQUIREMENTS:**
- Each match pair must be on ONE line
- Use three dashes with spaces: ` --- ` to separate premise from match
- NO empty lines between match pairs
- NO "Correct matches:" section needed
- Script parses the `---` separator to identify pairs

```
Matching

Match each item to its description

Premise 1 --- Match A
Premise 2 --- Match B
Premise 3 --- Match C

---
```

**Example Matching Question:**
```
Matching

Match each keyboard shortcut to its function.

Cmd/Ctrl + T --- Enter Free Transform mode
Shift + Alt/Option --- Scale proportionally from center
Shift while rotating --- Constrain to 15-degree increments

---
```

### Structure Guidelines:
- **Always start with section headers** before questions of that type
- Group all questions by type under their header
- Include setup instructions at top (question pool size, number delivered, points, attempts)
- No decorative formatting - keep it plain text for easy Canvas import
- Write clear, unambiguous questions
- Ensure distractors (wrong answers) are plausible
- Align questions with module learning outcomes
- Use separators (---) between questions for readability

### Setup Instructions Template:
```
Setup Instructions:
- Question pool: 8 questions
- Number delivered: 5 questions
- Points per question: 2 points
- Total possible: 10 points
- Attempts allowed: Unlimited
- Show correct answers: After submission
```

---

## RUBRIC FORMATTING INSTRUCTIONS

### Structure:
- Each criterion on its own section
- Blank lines between achievement levels for readability
- Level name and point range on **separate lines**
- Weight criteria by importance (higher points = more important)
- Keep top tier point range narrow (1-2 point spread)

### Achievement Level Options:

**For Creative/Design Assignments:**
- Exemplary
- Proficient
- Developing
- Incomplete/Missing

**For Task-Based/Technical Assignments:**
- Successfully Completed
- Mostly Successful
- Not Completed Successfully
- Not Completed

### Format Structure (Creative Assignment Example):
```
Criterion Name: [Description of what's being evaluated]

Exemplary
18-20 pts
Description of exemplary performance

Proficient
15-17 pts
Description of proficient performance

Developing
10-14 pts
Description of developing performance

Incomplete/Missing
0-9 pts
Description of incomplete or missing work

---
```

### Format Structure (Task-Based Assignment Example):
```
Criterion Name: [Description of what's being evaluated]

Successfully Completed
9-10 pts
Description of successful completion

Mostly Successful
7-8 pts
Description of mostly successful completion

Not Completed Successfully
4-6 pts
Description of unsuccessful completion

Not Completed
0-3 pts
Description of not completed work

---
```

### Rubric Guidelines:
- **Point values are variable** - adjust based on criterion importance
- **Number of criteria varies** by assignment complexity
- **Descriptions should be specific** to the criterion being evaluated
- Use triple dashes (---) as separators between criteria
- No decorative formatting - plain text for Canvas import
- Align criteria with assignment requirements and module outcomes
- Top tier should have narrow point range (encourages excellence)
- Make descriptions measurable and observable

---

## OUTPUT PREFERENCES

### File Formats:
- **Module Overviews**: HTML format with embedded styling
- **Assignments**: HTML format with embedded styling
- **Lesson Content**: Word documents (.docx) or Markdown, formatted for PowerPoint conversion using paragraph styles
- **Knowledge Checks/Quizzes**: Word documents (.docx) for QTI conversion
- **Rubrics**: Plain text format for easy Canvas copy/paste

### Workflow:
1. Draft content in Markdown artifacts for review and editing
2. Convert to final format (HTML or .docx) when approved
3. Maintain consistent formatting across all module materials
4. Ensure all materials align with module outcomes and assignments

---

## QUALITY STANDARDS

### Alignment (Quality Matters):
- Learning outcomes align with assessments
- Assessments align with instructional materials
- Instructional materials support achievement of outcomes
- Clear connection between all course components

### Accessibility:
- Clear, descriptive headings
- Logical content hierarchy
- Readable formatting
- Alternative text considerations for images (when applicable)

### Clarity:
- Unambiguous instructions
- Specific requirements and expectations
- Clear assessment criteria
- Professional, instructional tone


*These instructions should be applied consistently across all course module development unless specifically instructed otherwise for a particular module or assignment.*

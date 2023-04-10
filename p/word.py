import docx


document = docx.Document('testerDOC.docx')

word = 'link'

# Search for the word
for paragraph in document.paragraphs:
    paragraph_text = ''
    for run in paragraph.runs:
        if run.text:  # only include non-empty Run objects
            paragraph_text += run.text
            if word in run.text:
                print(f'Found the word "{word}" in run {run.text} with style {run.style}')
    print(f'Paragraph text: {paragraph_text}')
    


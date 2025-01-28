import os
import random
import string
from docx import Document

# Create a folder called 'test' if it doesn't already exist
if not os.path.exists('test'):
    os.makedirs('test')

for i in range(500):
    # Generate a random filename
    filename = os.path.join('test', ''.join(random.choices(string.ascii_lowercase, k=10)) + '.docx')
    print('Generating...')
    
    # Generate a random text content of size between 1-5 MB
    doc = Document()
    size_limit = random.randint(1, 5) * 10**6
    current_size = 0
    while current_size < size_limit:
        word_count = random.randint(50, 200)
        words = [''.join(random.choices(string.ascii_letters, k=random.randint(1, 10))) for _ in range(word_count)]
        
        # Generate a random paragraph of size between 1-5 KB
        # paragraph_size = random.randint(1, 5) * 10**3
        text = ' '.join(words)
        doc.add_paragraph(text)
        current_size += len(text.encode('utf-8'))
    
    # Save the document with the random filename
    doc.save(filename)
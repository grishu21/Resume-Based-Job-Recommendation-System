import spacy

# Load the pre-trained spaCy model
nlp = spacy.load("en_core_web_sm")

def extract_skills(text):
    # Predefined list of common skills (programming languages, tools, etc.)
    predefined_skills = ['Python', 'Java', 'C', 'C++', 'SQL', 'R', 'HTML', 'CSS', 'Machine Learning', 'TensorFlow', 
                         'Keras', 'Git', 'Flask', 'REST API', 'Data Science', 'Cybersecurity', 'AI', 'Artificial Intelligence',
                         'Communication', 'Teamwork', 'Data Structures', 'Algorithms', 'Excel', 'Photoshop', 'Illustrator', 'JavaScript']

    # Process the extracted text using spaCy's NLP pipeline
    doc = nlp(text)
    
    # Extract potential skills based on predefined skills list
    extracted_skills = set()

    # Check for programming languages, tools, or related skills
    for token in doc:
        if token.text.lower() in [skill.lower() for skill in predefined_skills]:
            extracted_skills.add(token.text)
    
    # You can also use Named Entity Recognition (NER) to extract entities related to technologies/tools
    for ent in doc.ents:
        if ent.label_ in ['ORG', 'GPE', 'PRODUCT']:  # e.g., product or tool names could be classified as ORG or PRODUCT
            extracted_skills.add(ent.text)
    
    return list(extracted_skills)

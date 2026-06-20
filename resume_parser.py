import spacy

def extract_skills_from_resume(text, skill_extractor):
    """
    Extract skills from resume text using SkillExtractor.
    Filters only unique skills.
    """
    annotations = skill_extractor.annotate(text)

    # Extract both full and partial matches
    skills = []
    if 'results' in annotations:
        if 'full_matches' in annotations['results']:
            skills.extend([item['doc_node_value'] for item in annotations['results']['full_matches']])
        if 'ngram_scored' in annotations['results']:
            skills.extend([item['doc_node_value'] for item in annotations['results']['ngram_scored']])

    # Remove duplicates and unwanted words like 'Manipal'
    cleaned_skills = list(set([s for s in skills if len(s) > 1 and s.isalpha()]))

    return cleaned_skills

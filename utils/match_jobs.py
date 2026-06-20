# Match extracted skills to job descriptions
def match_jobs(skills, job_list):
    matched_jobs = []
    for job in job_list:
        if any(skill.lower() in job['description'].lower() for skill in skills):
            matched_jobs.append(job)
    return matched_jobs

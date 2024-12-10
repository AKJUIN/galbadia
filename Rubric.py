from canvasapi import Canvas
import pandas as pd

# User configuration
API_URL = "https://<your-canvas-domain>.instructure.com"  # Replace with your Canvas domain
ACCESS_TOKEN = "your-access-token"  # Replace with your Canvas Access Token
SEARCH_CRITERION = "specific keyword or phrase"  # Replace with the criterion you're searching for

# Connect to Canvas
canvas = Canvas(API_URL, ACCESS_TOKEN)

# Fetch courses
def get_courses(canvas_instance):
    return canvas_instance.get_courses()

# Fetch rubrics for a course
def get_rubrics(course):
    rubrics_data = []

    try:
        rubrics = course.get_rubrics()

        for rubric in rubrics:
            criteria_found = False
            rubric_criteria = []

            # Iterate through rubric criteria
            for criterion in rubric.data.get("criteria", []):
                rubric_criteria.append({
                    "Criterion Description": criterion.get("description", ""),
                    "Points": criterion.get("points", "")
                })
                # Check if the criterion matches the search term
                if SEARCH_CRITERION.lower() in criterion.get("description", "").lower():
                    criteria_found = True

            rubrics_data.append({
                "Rubric Name": rubric.title,
                "Criteria Found": criteria_found,
                "Criteria Details": rubric_criteria,
                "Rubric Link": f"{API_URL}/courses/{course.id}/rubrics/{rubric.id}",
            })

    except Exception as e:
        print(f"Error fetching rubrics for course {course.name}: {e}")

    return rubrics_data

# Export to Excel
def export_to_excel(data, filename="canvas_rubric_search_results.xlsx"):
    writer = pd.ExcelWriter(filename, engine='openpyxl')

    for course_name, rubrics in data.items():
        df = pd.DataFrame(rubrics)
        df.to_excel(writer, sheet_name=course_name[:30], index=False)  # Sheet names max 31 chars

    writer.save()
    print(f"Data exported to {filename}")

if __name__ == "__main__":
    try:
        print("Fetching courses and rubrics from Canvas...")
        courses = get_courses(canvas)
        course_rubrics_data = {}

        for course in courses:
            print(f"Processing course: {course.name}")
            rubrics = get_rubrics(course)
            if rubrics:
                course_rubrics_data[course.name] = rubrics

        if course_rubrics_data:
            export_to_excel(course_rubrics_data)
            print("Export completed successfully!")
        else:
            print("No rubrics found or no data to export.")
    except Exception as e:
        print(f"An error occurred: {e}")

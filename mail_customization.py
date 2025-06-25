import csv
import os
from dotenv import load_dotenv
import google.generativeai as genai
import requests

# Load environment variables from .env file
load_dotenv()


def google_search(query, api_key, cse_id):
    url = "https://www.googleapis.com/customsearch/v1"
    params = {"q": query, "key": api_key, "cx": cse_id}
    response = requests.get(url, params=params)
    results = response.json()
    snippets = []
    for item in results.get("items", []):
        snippets.append(item.get("snippet", ""))
    return "\n".join(snippets[:3])  # Use top 3 snippets


def generate_mail_csv(input_csv, output_csv):
    consultant_name = input("Enter consultant name: ")
    consultant_email = input("Enter consultant email: ")

    # Configure the Gemini API key from environment variables
    gemini_api_key = os.getenv("GEMINI_API_KEY")
    if not gemini_api_key:
        print("Error: GEMINI_API_KEY not found in .env file.")
        return
    genai.configure(api_key=gemini_api_key)

    # Use a supported model name
    model = genai.GenerativeModel("models/gemini-2.5-pro")

    with open(input_csv, "r", encoding="utf-8") as f_in, open(
        output_csv, "w", newline="", encoding="utf-8"
    ) as f_out:
        reader = csv.DictReader(f_in)
        fieldnames = ["From", "To", "Subject", "Body"]
        writer = csv.DictWriter(f_out, fieldnames=fieldnames)
        writer.writeheader()

        for row in reader:
            sender_id = row.get("Email", "")
            company_name = row.get("Company_Name", "")
            company_info = row.get("About", "")
            company_keywords = row.get("Keywords", "")
            company_location = row.get("Location", "")
            company_industries = row.get("Industries", "")
            company_owner = row.get("Owner", "")

            # Fetch extra info from Google Search
            google_api_key = os.getenv("GOOGLE_API_KEY")
            google_cse_id = os.getenv("GOOGLE_CSE_ID")
            search_snippets = google_search(
                f"{company_name} about", google_api_key, google_cse_id
            )

            # Detailed system prompt
            prompt_text = (
                "You are an expert business consultant tasked with writing highly professional, "
                "personalized outreach emails to companies. Each email should include:\n"
                "- A compelling, relevant subject line.\n"
                "- A detailed, friendly, and professional body that references the company's background and mission.\n"
                "- The email should be from the consultant (details below) to the company (details below).\n"
                "- Do not include any labels like 'Subject:' or 'Body:'.\n"
                "- The output must be the subject line, then a newline, then the full email body.\n"
                "- The email should be suitable for a first contact and encourage a reply.\n"
                "- The length limit is two paragraphs, be detailed on what 180DC IITKGP can offer to them and be direct.\n"
                "Here is a sample mail, you must refer to a similar format only for sending the mails"
                """Respected Sir,

                    I am {consultant_name}, a Consultant at 180 Degrees Consulting, IIT Kharagpur. Thank you for your response on LinkedIn. We're a student-run body that is passionate about providing operational and strategic services to NGOs and social enterprises, and be a part of their growth and impact journey.
                    We have been greatly influenced by the powerful contributions of Terre des hommes foundation in protecting children's rights and well-being and provide vital protection through core programs. Your efforts have profoundly impacted the lives of countless individuals through the events you organize and the awareness you generate.

                    We at 180DC IIT Kharagpur have worked with organizations such as the CRY Foundation and Robin Hood Army, aiding them in improving operations, strategic growth, and program outcomes. We feel our services may assist Terre des hommes cause by multiplying its impact through strategy and operational consulting.

                    We would be delighted to reach out to you and learn about how we can work together. To provide more insights about our organisation I have attached our brochure. We look forward to speaking with you and your team!

                    Best regards,
                    {consultant_name}
                    Consultant
                    180 Degrees Consulting, IIT Kharagpur
                    https://www.180dc.org/branches/IITKGP 
                """
                "This must be modified according to the purpose and nature of the company, you may perform google search and online research to find more information about the company."
                f"Consultant Name: {consultant_name}\n"
                f"Consultant Email: {consultant_email}\n"
                f"Company Name: {company_name}\n"
                f"Company Info: {company_info}\n"
                f"Company Keywords: {company_keywords}\n"
                f"Company Location: {company_location}\n"
                f"Company Industries: {company_industries}\n"
                f"Company Owner: {company_owner}\n"
                f"Recent Google Search Results: {search_snippets}\n"
            )

            try:
                response = model.generate_content(prompt_text)
                generated_text = (
                    response.text
                    if hasattr(response, "text")
                    else (
                        response.generations[0].content if response.generations else ""
                    )
                )
                text_output = generated_text.strip().split("\n", 1)
                subject_line = text_output[0] if text_output else "Following Up"
                body_text = (
                    text_output[1]
                    if len(text_output) > 1
                    else "Could not generate email body."
                )
            except Exception as e:
                print(f"Error generating content for {company_name}: {e}")
                subject_line = f"Introduction from {consultant_name}"
                body_text = f"Dear {company_name} team,\n\nI am writing to introduce our consulting services..."

            writer.writerow(
                {
                    "From": consultant_email,
                    "To": sender_id,
                    "Subject": subject_line,
                    "Body": body_text,
                }
            )
        print(f"Successfully generated mails in {output_csv}")


if __name__ == "__main__":
    input_csv_path = r"c:\Work\College\Coding\CNS_Automation\company_list.csv"
    output_csv_path = r"c:\Work\College\Coding\CNS_Automation\generated_mails.csv"
    generate_mail_csv(input_csv_path, output_csv_path)

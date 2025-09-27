# Weekly Report Generator

This project automates the creation of 24 weekly intern reports. It uses a sophisticated workflow involving Google's AI Studio to generate structured JSON data from raw notes, which is then used to populate a DOCX template. Finally, all generated reports are compiled into a single PDF document.

## Features

- **AI-Powered Data Generation**: Uses Google AI Studio with the Gemini 2.5 Pro model to process raw notes and generate detailed, structured weekly data in JSON format.
- **Automated Document Filling**: A Python script reads the generated JSON data and populates a `Daily Report Template.docx`.
- **Signature Integration**: Automatically embeds a predefined signature image (`signature.png`) into each report.
- **PDF Compilation**: Converts all the generated DOCX reports into a single, combined PDF file for easy sharing and archiving.
- **Preserves Formatting**: Maintains the original formatting of the DOCX template in the final output.

## Files

- `fill_weekly_reports.py`: The main script to fill the DOCX template with data from the JSON file.
- `convert_docx_to_pdf.py`: A script to convert the generated DOCX files into a single PDF.
- `data/weekly_data.json`: The file where you will save the generated JSON data from Google AI Studio.
- `data/Daily Report Template.docx`: The MS Word template for the weekly report.
- `data/signature.png`: Your signature image file.
- `requirements.txt`: A file listing the Python dependencies.

## Workflow

Follow these steps to generate your combined weekly report PDF:

### Step 1: Generate Weekly Data with Google AI Studio

1.  **Open Google AI Studio**: Navigate to [https://aistudio.google.com/prompts/new_chat](https://aistudio.google.com/prompts/new_chat).

2.  **Select the Model**: At the top of the left sidebar, choose the model `models/gemini-2.5-pro`.

3.  **Enable Structured Output**: Turn on the **Structured Output** toggle button.

4.  **Run Prompts**: Copy and paste the following prompts one by one into the input box and click the **Generate** button for each.

    **Prompt 1:**
    ```
    Give me all the 24 weeks ending dates from [Intern Start Date] to [Ending Date]. 2025 have 28 days in feb
    ```

    **Prompt 2:**
    ```
    Here are the full note I have created while working on my training canvas from [Intern Start Date] to [Ending Date]:
    """
    [Paste your full notes here]
    """

    Now I need to create weekly reports for each week ending on the above dates. Here is the sample JSON structure I need:
    {
    "week_no": "02",
    "week_ending": "Sunday: 28-07-2025",
    "training_mode": "Remote",
    "weekly_activities": [
    {"day": "MONDAY", "date": "22-07", "description": "Morning safety briefing. Site inspection and progress review. Equipment check and calibration."},
    {"day": "TUESDAY", "date": "23-07", "description": "Concrete pouring for foundation Section A. Quality control testing. Supervision of reinforcement work."},
    {"day": "WEDNESDAY", "date": "24-07", "description": "Laboratory testing of concrete samples. Material quality assessment. Documentation of test results."},
    {"day": "THURSDAY", "date": "25-07", "description": "Structural steel installation supervision. Welding quality inspection. Safety compliance audit."},
    {"day": "FRIDAY", "date": "26-07", "description": "Weekly progress meeting. Site measurement and surveying. Preparation of progress reports."},
    {"day": "SATURDAY", "date": "27-07", "description": "Equipment maintenance and calibration. Site cleanup activities. Material inventory check."},
    {"day": "SUNDAY", "date": "28-07", "description": "Rest day. Review of weekly activities and preparation for next week's tasks."}
    ],
    "details_notes": "This week was highly productive with significant progress in foundation work. All concrete pours were completed successfully with quality test results exceeding specifications. The structural steel installation commenced on schedule. Weather conditions remained favorable throughout the week with no delays due to rain. All safety protocols were strictly followed with zero incidents reported. The laboratory testing revealed excellent concrete strength results. Equipment performed well with only routine maintenance required. Team coordination was excellent and all deadlines were met.",
    "engineer_remarks": "Excellent progress demonstrated this week. The trainee showed strong technical competency in concrete quality control and steel installation supervision. Good understanding of safety protocols and documentation procedures. Recommend continuing with advanced structural work supervision next week. Overall performance is above expectations.",
    "engineer_date": "28-07-2025",
    "engineer_designation_signature": "Mr.Ben Basuni, - Chief Technology Officer(CTO)"
    }

    Use the week_ending date for the engineer_date. Write all other fields (week_no, week_ending, weekly_activities, details_notes, engineer_remarks) based on my tasks and dates. If no task is found for a specific date in my notes, leave the description as an empty string.

    Give me the JSON for Week 01.
    ```

    **Prompt 3:**
    ```
    now give me all 24 weeks in a list structure like this:

    [
    {
    "week_no": "01",
    "week_ending": "Sunday: 23-02-2025",
    "training_mode": "Remote",
    "weekly_activities": [
    {
    "day": "MONDAY",
    "date": "17-02",
    "description": "Onboarding meetings with the team. Cloned and explored the 'dabu-be-worker' codebase to get familiar with existing projects. Learned about RunPod for GPU deployment."
    },
    {
    "day": "TUESDAY",
    "date": "18-02",
    "description": "Conducted self-study on cloud GPUs and serverless architecture. Set up a dual-boot Ubuntu development environment, resolving installation and configuration issues."
    },
    {
    "day": "WEDNESDAY",
    "date": "19-02",
    "description": "Successfully ran the DABU backend worker. Learned advanced Git commands and workflows for branching and merging. Integrated recent changes from the dev branch."
    },
    {
    "day": "THURSDAY",
    "date": "20-02",
    "description": "Began work on the new PhonicsMaker project. Created a detailed implementation plan, researched key technologies like WeasyPrint and DALL-E, and finalized the system architecture."
    },
    {
    "day": "FRIDAY",
    "date": "21-02",
    "description": "Initialized the PhonicsMaker backend repository. Began development by debugging the base template and started integrating core project features."
    },
    {
    "day": "SATURDAY",
    "date": "22-02",
    "description": ""
    },
    {
    "day": "SUNDAY",
    "date": "23-02",
    "description": ""
    }
    ],
    "details_notes": "This was an intensive first week focused on onboarding and setting up for a new major project. I quickly familiarized myself with the company's existing codebase and development tools like RunPod. A significant effort was dedicated to establishing a stable Ubuntu development environment. By mid-week, I was assigned to the new PhonicsMaker project, where I created a comprehensive implementation plan and researched the core technologies. The week concluded with the successful initialization of the project repository and the start of active development, including debugging the foundational template.",
    "engineer_remarks": "Tharindu had an excellent start. He demonstrated a proactive attitude by quickly setting up his environment and familiarizing himself with our projects. His detailed planning and research for the new PhonicsMaker project were impressive and set a strong foundation for its success. He is a fast learner and has integrated well with the team.",
    "engineer_date": "23-02-2025",
    "engineer_designation_signature": "Mr.Ben Basuni, - Chief Technology Officer(CTO)"
    },
    ...
    ]
    ```

5.  **Save the JSON Data**: Copy the complete JSON list generated by the final prompt and paste it into the `data/weekly_data.json` file, replacing any existing content.

### Step 2: Set Up the Environment

1.  **Create a virtual environment**:
    ```cmd
    py -m venv .venv
    ```

2.  **Activate the environment**:
    ```cmd
    .venv\Scripts\activate
    ```

3.  **Install the required dependencies**:
    ```cmd
    pip install -r requirements.txt
    ```

### Step 3: Run the Scripts

1.  **Generate the Weekly Reports**:
    Run the `fill_weekly_reports.py` script to create the individual `.docx` files in the `Weekly Reports` folder.
    ```cmd
    python fill_weekly_reports.py
    ```

2.  **Convert Reports to PDF**:
    Run the `convert_docx_to_pdf.py` script to combine all the generated reports into a single PDF.
    ```cmd
    python convert_docx_to_pdf.py
    ```

## Output

After running the scripts, you will find:
- Individual weekly reports in `.docx` format inside the `Weekly Reports/` folder.
- A final combined PDF document named `Combined_Weekly_Reports.pdf` in the root directory.

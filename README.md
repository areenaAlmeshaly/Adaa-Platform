# Adaa | أداء
### Making Employee Contributions More Visible

Adaa explores how digital work activity can provide a clearer picture of employee contributions and support more informed performance discussions.

Developed as an MVP during **Tuwaiq Hackathon**.

## The Problem

Employee performance evaluations often rely on predefined goals or subjective assessments. However, these methods may not capture the full picture of an employee's daily contributions, especially in increasingly digital work environments.

This raises a question:

**How can everyday digital work provide additional context for understanding employee contributions?**

## The Approach

Adaa explores the use of selected digital activity indicators to make parts of an employee's daily work more visible.

Rather than replacing human judgment, the platform aims to provide additional information that supervisors can consider during performance discussions.

## The Prototype

The MVP currently includes:

- **Activity Tracking:** Tracks selected work activity in Microsoft Excel on macOS.
- **Data Analysis:** Processes task-related metadata and activity indicators.
- **AI-Assisted Insights:** Uses Gemini API to generate summaries from collected data.
- **Dashboard:** Presents activity indicators and generated insights.
- **Activity History:** Stores historical records locally.

## How It Works

`Activity Tracking → Data Processing → AI Analysis → Dashboard`

## Screenshots


![Adaa Platform Interface](docs/images/adaa-platform-interface.png)

![Adaa Platform Overview](docs/images/adaa-platform-overview.png)


## Technology Stack

Python · FastAPI · Gemini API · AppleScript · JavaScript · HTML/CSS · JSON

## Getting Started

**Requirements:** macOS, Python 3.10+, Microsoft Excel, and a Gemini API key.

```bash
git clone https://github.com/areenaAlmeshaly/Adaa-Platform.git
cd Adaa-Platform

python3 -m venv .venv
source .venv/bin/activate

pip install -r requirements.txt
```

Create a `.env` file and add your Gemini API key:

```env
GEMINI_API_KEY=your_api_key_here
```

Start the application:

```bash
python start.py
```

Open `http://127.0.0.1:8000` in your browser.

## Privacy & Limitations

Adaa does not capture screenshots or keystrokes. Activity tracking and dashboard data are handled locally, while selected metadata is sent to the Gemini API for analysis.

The current MVP does not include comprehensive anonymization, encryption, or access controls. Activity indicators should not be interpreted as definitive measures of productivity or performance. Some dashboard views use sample data.

## Project Status

Adaa is an early-stage prototype developed during Tuwaiq Hackathon using AI-assisted development. Its indicators, assumptions, and real-world usefulness require further testing and validation.




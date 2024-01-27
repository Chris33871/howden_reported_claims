# Howden Reported Claims

This repository is the completion of the Howden Group technology assessment.

There are 4 scrips to run in order to get the results to each question. The PDF with the questions is not included for privacy concerns. That along with the API key needed to run one of the scripts will be shared via email.

## Table of Contents

- [Table of Contents](#table-of-contents)
- [Features](#features)
- [Prerequisites](#prerequisites)
- [Getting Started](#getting-started)

  - [Installation](#installation)
  - [Run](#Running-the-Project)

- [Footer](#Footer)

## Features

- Creating sql insert scripts
- Creating visualisations
- Getting currency exchange rates

## Prerequisites

- Python (version 3)
- [Other dependencies or prerequisites]

## Getting Started

Provide step-by-step instructions on how to set up the project locally. Include any necessary installation or configuration steps.

### Prerequisites

- Python 3.10

### Installation

1. Clone this repository:

   ```bash
   git clone https://github.com/Chris33871/howden_reported_claims.git
   ```

2. Change into the project directory:

   ```bash
   cd howden_reported_claims
   ```

3. Create a virtual environment:

   ```bash
   python -m venv venv
   ```

   or `Virtualenv venv`

4. Activate the virtual environment:

   - On Windows:

     ```bash
     venv\Scripts\activate
     ```

   - On macOS and Linux:

     ```bash
     source venv/bin/activate
     ```

5. Install project dependencies:

   ```bash
   pip install -r requirements.txt
   ```

### Running the Project

Create a .env file and follow the .env.template file. The API key will be shared by mail.

To run a script, execute
`python utils/name_of_the_script_to_run.py`

The script will do as its name indicates

Keep in mind that there is a limited amount of API request quota available and only 2 weeks of free usage.

API request left since last run: 1300 - Free days left since last run: 13

##### Footer

API used: [ExchangeRate-API](https://app.exchangerate-api.com/dashboard)

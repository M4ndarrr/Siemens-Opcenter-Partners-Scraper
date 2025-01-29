# Siemens Opcenter Partners Scraper

A Python tool to parse HTML input from the Siemens Opcenter Partners platform and extract structured information into an Excel file for easy analysis.

[Explore the Documentation »](#)  

[Report a Bug](#) · [Request a Feature](#)

---

## Table of Contents  

- [About the Project](#about-the-project)  
- [Getting Started](#getting-started)  
  - [Prerequisites](#prerequisites)  
  - [Installation](#installation)  
- [Usage](#usage)  
- [License](#license)  
- [Contact](#contact)  

---

## About the Project  

The **Siemens Opcenter Partners Scraper** automates the extraction of partner-related data from HTML files originating from the Siemens Opcenter platform. Instead of manually sifting through HTML, this tool parses the document and exports the information into a clean Excel format, streamlining analysis and reporting.  

### Key Features  

✅ Parses HTML files exported from Siemens Opcenter Partners platform  
✅ Extracts relevant partner and data fields  
✅ Saves structured information in an Excel file  
✅ Simple and efficient automation  

[Back to top](#table-of-contents)

---

## Getting Started  

### Prerequisites  

1. **Python 3.7+** – [Download here](https://www.python.org/downloads/)  
2. **Required Libraries** – Install dependencies using:  

   ```bash
   pip install requests beautifulsoup4 pandas openpyxl
   ```

### Installation  

1. **Clone the repository**  

   ```bash
   git clone https://github.com/M4ndarrr/Siemens-Opcenter-Partners-Scraper.git
   cd Siemens-Opcenter-Partners-Scraper
   ```

2. **Install dependencies**  

   ```bash
   pip install -r requirements.txt
   ```

3. **Run the scraper**  

   ```bash
   python main.py
   ```

[Back to top](#table-of-contents)

---

## Usage  

### Prepare your HTML file  

Export or save the Siemens Opcenter Partners HTML page to your local system.  

### Run the script  

```bash
python main.py
```  

### Output  

The extracted data will be stored in `output` in the project directory.  

[Back to top](#table-of-contents)

---

## License  

Distributed under the MIT License. See `LICENSE.txt` for more details.  

[Back to top](#table-of-contents)

---

## Contact  

Jan Tichy  
Email: jan.tichy@jnt-digital.net  

[Back to top](#table-of-contents)

📊 Superstore Sales Dashboard
This project is a Flask web application that reads data from Superstore.xlsx and generates an interactive sales analytics dashboard. It calculates yearly sales, profits, quantities, discounts, and shipment modes, along with state-wise summaries, and displays them using an HTML template (SalesDashboard.html).

🚀 Features
Product & Customer Analytics

Total number of unique products

Total number of unique customers

Year-Based Insights

Available sales years

Yearly Quantity sold

Yearly Sales totals

Yearly Profit/Loss totals

Yearly Discount totals

Shipment Mode Analysis

Year-wise shipment mode counts

State-wise Insights

Year-wise & state-wise Sales, Profit, and Quantity

Interactive Web Dashboard

Renders all processed data into a dynamic HTML dashboard

🛠️ Tech Stack
Python 3

Flask — Web framework

openpyxl — Excel file reading

Jinja2 — Templating engine for HTML

HTML/CSS — Frontend UI

📂 Project Structure
text
.
├── app.py                             # Main Flask application (your code)
├── Superstore.xlsx                    # Input dataset
├── templates/
│   └── SalesDashboard.html            # Dashboard HTML template
├── static/                            # (Optional) CSS/JS files
└── README.md                          # Project documentation
⚙️ Installation & Setup
1️⃣ Clone the Repository

bash
git clone https://github.com/yourusername/superstore-dashboard.git
cd superstore-dashboard
2️⃣ Install Dependencies

bash
pip install flask openpyxl
3️⃣ Place Your Data File

Ensure your Superstore.xlsx is in the same directory as app.py

The Excel file must have the following columns:

Order Date

Quantity

Sales

Profit

Ship Mode

State

Discount

Product ID / Customer ID (mapped to column Q and column G in the given script)

4️⃣ Run the Application

bash
python app.py
5️⃣ Access the Dashboard

Open your browser and go to:

text
http://127.0.0.1:5000/
📊 Output Example
The dashboard will display:

Total Products: XXXX

Total Customers: YYYY

Years Available: 2015, 2016, 2017, 2018

Profit/Loss Summary

Year-wise Sales Graph

Year-wise Quantity Bar Chart

Shipment Mode Summary Table

State-wise Yearly Performance Table

📝 Notes
The provided Superstore.xlsx must follow the Superstore Sample Dataset format.

SalesDashboard.html must be located inside the templates folder.

You can customize the visuals in the HTML template to include charts using Chart.js, Plotly, or any other library.

Debug mode is enabled for development; disable it before deployment.

📄 License
This project is open-source. You may modify and distribute it under your own terms.


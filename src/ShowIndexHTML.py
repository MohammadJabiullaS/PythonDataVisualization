from flask import Flask, render_template

app = Flask(__name__, template_folder='../templates')
@app.route('/')
def index():
    # Sample data to pass to the HTML template
    user_data = {"name": "Alice", "age": 30}
    return render_template('SalesDashboard.html', user=user_data)

if __name__ == '__main__':
    app.run(debug=True)
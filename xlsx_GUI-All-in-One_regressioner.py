import sys
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QFileDialog, QComboBox

class PolynomialRegressionApp(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Xlsx Regression Graph Generator")
        self.setGeometry(100, 100, 400, 350)

        self.initUI()

    def initUI(self):
        # Create input fields and labels
        self.title_label = QLabel("Graph Title:")
        self.title_input = QLineEdit()

        self.degree_label = QLabel("Polynomial Degree:")
        self.degree_combo = QComboBox()
        self.degree_combo.addItems(["1 (Linear)", "2 (Quadratic)", "3 (Cubic)", "4", "5", "6", "7", "8", "9", "10"])

        self.x_label = QLabel("X-axis Label:")
        self.x_input = QLineEdit()

        self.y_label = QLabel("Y-axis Label:")
        self.y_input = QLineEdit()

        self.xcol_label = QLabel("X-axis Column:")
        self.xcol = QLineEdit()

        self.ycol_label = QLabel("Y-axis Column:")
        self.ycol = QLineEdit()

        self.file_label = QLabel("File Name:")
        self.file_input = QLineEdit()
        self.file_button = QPushButton("Select File")
        self.file_button.clicked.connect(self.select_file)

        self.generate_button = QPushButton("Generate Graph")
        self.generate_button.clicked.connect(self.generate_graph)

        # Layout for the inputs and button
        layout = QVBoxLayout()
        layout.addWidget(self.title_label)
        layout.addWidget(self.title_input)
        layout.addWidget(self.degree_label)
        layout.addWidget(self.degree_combo)
        layout.addWidget(self.x_label)
        layout.addWidget(self.x_input)
        layout.addWidget(self.y_label)
        layout.addWidget(self.y_input)
        layout.addWidget(self.xcol_label)
        layout.addWidget(self.xcol)
        layout.addWidget(self.ycol_label)
        layout.addWidget(self.ycol)
        layout.addWidget(self.file_label)
        layout.addWidget(self.file_input)
        layout.addWidget(self.file_button)
        layout.addWidget(self.generate_button)
        self.setLayout(layout)

    def select_file(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_name, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
        if file_name:
            self.file_input.setText(file_name)

    def generate_graph(self):
        # Read data from Excel file
        filename = self.file_input.text()
        try:
            df = pd.read_excel(filename, engine='openpyxl')
        except FileNotFoundError:
            print(f"File not found: {filename}")
            return

        # Extract x and y values from the DataFrame
        x = df[self.xcol.text()].values
        y = df[self.ycol.text()].values
        x = x[~np.isnan(x)]
        y = y[~np.isnan(y)]
        # Get the selected degree from the combo box
        selected_degree = int(self.degree_combo.currentText().split()[0])

        # Perform polynomial regression
        coeffs = np.polyfit(x, y, selected_degree)
        mymodel = np.poly1d(coeffs)
        coeffs = mymodel.coefficients
        function_text = "f(x) = "
        for i, coef in enumerate(coeffs):
            if i == 0:
                function_text += f"{coef:.2f}x^{len(coeffs)-i-1}"
            else:
                function_text += f" + {coef:.2f}x^{len(coeffs)-i-1}" if coef >= 0 else f" - {-coef:.2f}x^{len(coeffs)-i-1}"

        # Shorten the last term, e.g., "x^0" to a constant value
        function_text = function_text.replace("x^0", "")
        plt.text(0.7, 0.05, function_text, horizontalalignment='center', verticalalignment='center', transform=plt.gca().transAxes, bbox=dict(facecolor='white', alpha=0.5))
        # Plot the data and the regression line
        myline = np.linspace(min(x), max(x), 100)
        plt.scatter(x, y, label='Data')
        plt.plot(myline, mymodel(myline), color='red', label='Function fit') 

        graph_title = self.title_input.text()
        x_label = self.x_input.text()
        y_label = self.y_input.text()

        plt.xlabel(x_label)
        plt.ylabel(y_label)
        plt.title(graph_title)
        plt.legend()
        plt.grid(True)
        plt.show()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = PolynomialRegressionApp()
    window.show()
    sys.exit(app.exec_())
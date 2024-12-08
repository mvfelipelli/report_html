import pandas as pd
import plotly.express as px
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib
from datetime import datetime
import webbrowser
import os

class ReportGenerator:
    def __init__(self):
        """Initialize the ReportGenerator class

        This class contains methods to generate a HTML report from a dictionary
        of DataFrames. The report is structured with sections, each containing
        a graph and a table. The graph and table are generated using plotly
        and pandas respectively.
        """
        self.style = """
        <!--
        This CSS code is used to style the HTML report
        -->
        <style>
            body {
                font-family: Arial, sans-serif;
                margin: 40px;
                color: #333;
            }
            .container {
                max-width: 1200px;
                margin: 0 auto;
            }
            h1, h2 {
                color: #2c3e50;
                padding-bottom: 10px;
                border-bottom: 2px solid #eee;
            }
            .graph-container {
                margin: 20px 0;
                padding: 15px;
                border: 1px solid #ddd;
                border-radius: 5px;
            }
            table {
                width: 100%;
                border-collapse: collapse;
                margin: 20px 0;
            }
            th, td {
                padding: 12px;
                border: 1px solid #ddd;
                text-align: left;
            }
            th {
                background-color: #f5f5f5;
            }
            tr:nth-child(even) {
                background-color: #f9f9f9;
            }
        </style>
        """
    
    def create_graph(self, data, x_col, y_col, title, graph_type='line'):
        """Create a graph using plotly

        Parameters:
        data: DataFrame containing the data to be plotted
        x_col: Name of the column containing the x-axis data
        y_col: Name of the column containing the y-axis data
        title: Title of the graph
        graph_type: Type of graph to be created (line, bar, etc.)

        Returns:
        HTML code for the graph
        """
        if graph_type == 'line':
            fig = px.line(data, x=x_col, y=y_col, title=title)
        elif graph_type == 'bar':
            fig = px.bar(data, x=x_col, y=y_col, title=title)
        
        # Customize the layout of the graph
        fig.update_layout(
            # Center the title
            title_x=0.5,
            # Add margins around the graph
            margin=dict(t=50, l=20, r=20, b=20),
            # Remove the background color
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)'
        )
        
        # Convert the figure to HTML code
        return fig.to_html(full_html=False, include_plotlyjs=True)

    def generate_report(self, data_dict, output_path='report.html'):
        """
        Generate HTML report and save it to a file

        Parameters:
        data_dict: Dictionary containing DataFrames and their descriptions
        output_path: Path where the HTML file will be saved
        """
        
        html_content = f"""
        <html>
        <head>
            {self.style}
        </head>
        <body>
            <div class="container">
                <h1>Business Performance Report</h1>
                <p>Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M')}</p>
        """
        
        # Add each section from the data dictionary
        for section_name, section_data in data_dict.items():
            df = section_data['data']
            graph_type = section_data.get('graph_type', 'line')
            x_col = section_data.get('x_col')
            y_col = section_data.get('y_col')
            
            html_content += f"""
                <div class="section">
                    <h2>{section_name}</h2>
            """
            
            # Add graph if x_col and y_col are specified
            if x_col and y_col:
                html_content += f"""
                    <div class="graph-container">
                        {self.create_graph(df, x_col, y_col, section_name, graph_type)}
                    </div>
                """
            
            # Add table
            html_content += f"""
                    <div class="table-container">
                        {df.to_html(index=False, classes='table')}
                    </div>
                </div>
            """
        
        html_content += """
            </div>
        </body>
        </html>
        """
        
        # Save the report
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        return output_path

class ReportSender:
    @staticmethod
    def send_report(sender_email, sender_password, recipient_email, subject, report_path):
        """
        Send the generated HTML report via email
        """
        # Read the HTML content
        with open(report_path, 'r', encoding='utf-8') as f:
            html_content = f.read()
        
        # Create email message
        msg = MIMEMultipart('alternative')
        msg['Subject'] = subject
        msg['From'] = sender_email
        msg['To'] = recipient_email
        
        # Attach HTML content
        msg.attach(MIMEText(html_content, 'html'))
        
        # Send email using Gmail's SMTP server
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(sender_email, sender_password)
            server.send_message(msg)

# Example usage
if __name__ == "__main__":
    # Sample data
    sales_data = pd.DataFrame({
        'Date': pd.date_range(start='2024-01-01', periods=6, freq='M'),
        'Revenue': [45000, 52000, 48000, 51000, 54000, 58000],
        'Units': [150, 170, 160, 165, 180, 190]
    })
    
    inventory_data = pd.DataFrame({
        'Product': ['Widget A', 'Widget B', 'Widget C'],
        'Stock': [150, 200, 180],
        'Reorder_Point': [100, 150, 120]
    })
    
    # Prepare data dictionary
    data_dict = {
        'Sales Performance': {
            'data': sales_data,
            'graph_type': 'line',
            'x_col': 'Date',
            'y_col': 'Revenue'
        },
        'Inventory Status': {
            'data': inventory_data,
            'graph_type': 'bar',
            'x_col': 'Product',
            'y_col': 'Stock'
        }
    }
    
    # Step 1: Generate report
    generator = ReportGenerator()
    report_path = generator.generate_report(data_dict, 'business_report.html')
    
    # # Open report in browser for validation
    # webbrowser.open('file://' + os.path.abspath(report_path))
    
    # # Step 2: Send report (after validation)
    # user_input = input("Would you like to send the report? (yes/no): ")
    
    # if user_input.lower() == 'yes':
    #     sender = ReportSender()
    #     sender.send_report(
    #         sender_email="your_email@gmail.com",
    #         sender_password="your_app_password",  # Use app-specific password for Gmail
    #         recipient_email="recipient@example.com",
    #         subject="Business Performance Report",
    #         report_path=report_path
    #     )
    #     print("Report sent successfully!")
    # else:
    #     print("Report sending cancelled.")
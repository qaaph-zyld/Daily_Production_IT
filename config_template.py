"""
Configuration Template for Production Dashboard
Copy this file to config.py and update with your actual database credentials
"""

# Database Configuration
DB_CONFIG = {
    # SQL Server name (e.g., 'localhost\\SQLEXPRESS', 'server.domain.com', or IP address)
    'server': 'YOUR_SQL_SERVER_NAME_HERE',
    
    # Database name
    'database': 'QADEE2798',
    
    # Authentication method
    # For Windows Authentication, use:
    'trusted_connection': 'yes',
    
    # For SQL Server Authentication, comment out 'trusted_connection' above and uncomment below:
    # 'username': 'your_username',
    # 'password': 'your_password',
}

# Server Configuration
SERVER_CONFIG = {
    'host': '0.0.0.0',  # Listen on all interfaces
    'port': 5000,       # Port number
    'debug': True,      # Set to False in production
}

# Dashboard Configuration
DASHBOARD_CONFIG = {
    'refresh_interval_minutes': 15,  # Auto-refresh interval in minutes
    'page_title': 'Adient Production Dashboard - Hourly Production IT',
}

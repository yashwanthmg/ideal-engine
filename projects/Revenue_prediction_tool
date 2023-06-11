import pandas as pd

# Step 1: Load the Data
# Assuming you have two Excel files named "orders1.xlsx" and "orders2.xlsx"
excel_files = ["orders1.xlsx", "orders2.xlsx"]
dfs = []  # List to store DataFrames

# Load Excel files into DataFrames
for file in excel_files:
    df = pd.read_excel(file)
    dfs.append(df)

# Merge DataFrames if needed
merged_df = pd.concat(dfs)

# Step 2: Data Preparation
# Perform any necessary data cleaning operations
cleaned_df = merged_df.dropna()  # Example: Remove rows with missing values

# Step 3: Data Analysis
# Identify regular ordering accounts
account_orders = cleaned_df.groupby("business_account_id").size()

# Calculate the probability of reordering
reorder_prob = account_orders / account_orders.sum()

# Determine potential order value increase
category_order_value = cleaned_df.groupby("gl_product_group_desc")["ops_without_tax"].sum()
potential_increase_category = category_order_value.idxmax()

# Step 4: Prediction Modeling (Example: Linear Regression)
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_squared_error

# Split the data into training and testing sets
X = cleaned_df[["ops_without_tax", "units"]]  # Features for prediction
y = cleaned_df["ops_without_tax"]  # Target variable

X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

# Build and train the model
model = LinearRegression()
model.fit(X_train, y_train)

# Make predictions on the test set
y_pred = model.predict(X_test)

# Evaluate the model
mse = mean_squared_error(y_test, y_pred)

# Step 5: Generate Output
# Generate the desired output based on the analysis and predictions
regular_accounts = reorder_prob[reorder_prob > 0.5].index.tolist()
potential_increase_category

# Step 6: Test and Refine
# Perform testing and refinement as necessary

# Print the output
output_data = {
    "Regular Ordering Accounts": regular_accounts,
    "Potential Category for Order Value Increase": [potential_increase_category],
    "Mean Squared Error": [mse]
}

output_df = pd.DataFrame(output_data)

output_filename = "output.xlsx"
output_df.to_excel(output_filename, index=False)

print("Output saved as:", output_filename)

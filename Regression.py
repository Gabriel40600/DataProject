import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LinearRegression, Ridge
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
import numpy as np

# Load the processed LinkedIn dataset
file_path = "processed_linkedin_data.xlsx"
df = pd.read_excel(file_path, sheet_name="Sheet1")

# Feature Engineering
# Creating new features based on text length, presence of hashtags, and sentiment (if available)
df['Post Length'] = df['Post content'].astype(str).apply(len)
df['Hashtag Count'] = df['Contains Hashtag']

def normalize_column(column):
    return (column - column.mean()) / column.std()

df['Normalized Length'] = normalize_column(df['Post Length'])
df['Normalized Hashtags'] = normalize_column(df['Hashtag Count'])

df = df[['Normalized Length', 'Normalized Hashtags', 'reactions']].dropna()

# Splitting data into training and test sets
X = df[['Normalized Length', 'Normalized Hashtags']]
y = df['reactions']
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

# Applying Linear Regression
linear_regressor = LinearRegression()
linear_regressor.fit(X_train, y_train)
y_pred_linear = linear_regressor.predict(X_test)

# Applying Ridge Regression
ridge_regressor = Ridge(alpha=1.0)
ridge_regressor.fit(X_train, y_train)
y_pred_ridge = ridge_regressor.predict(X_test)

# Evaluating models
def evaluate_model(y_true, y_pred, model_name):
    print(f"{model_name} Performance:")
    print(f"MAE: {mean_absolute_error(y_true, y_pred):.2f}")
    print(f"MSE: {mean_squared_error(y_true, y_pred):.2f}")
    print(f"RMSE: {np.sqrt(mean_squared_error(y_true, y_pred)):.2f}")
    print(f"R² Score: {r2_score(y_true, y_pred):.2f}\n")

evaluate_model(y_test, y_pred_linear, "Linear Regression")
evaluate_model(y_test, y_pred_ridge, "Ridge Regression")

# Visualizing results
plt.figure(figsize=(8, 5))
plt.scatter(y_test, y_pred_linear, label="Linear Regression", alpha=0.6)
plt.scatter(y_test, y_pred_ridge, label="Ridge Regression", alpha=0.6, marker='x')
plt.plot([y.min(), y.max()], [y.min(), y.max()], linestyle="--", color='red')
plt.xlabel("Actual Engagement (Reactions)")
plt.ylabel("Predicted Engagement")
plt.title("Actual vs Predicted Engagement")
plt.legend()
plt.show()

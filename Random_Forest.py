import pandas as pd
import numpy as np
from scipy import stats
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.ensemble import RandomForestRegressor
import requests
import openpyxl
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import os
from sklearn.impute import SimpleImputer
from sklearn.metrics import mean_absolute_error, mean_squared_error
from math import sqrt

def download_gdp_data(api_indicator, start_year, end_year, file_path):
    # Định nghĩa URL API
    base_url = f"http://api.worldbank.org/v2/country/all/indicator/{api_indicator}?date={start_year}:{end_year}&format=json&per_page=1000"

    try:
        # Gửi yêu cầu GET đến API của Ngân hàng Thế giới và chuyển đổi phản hồi thành JSON
        response = requests.get(base_url)
        response.raise_for_status()  # Kiểm tra lỗi trong phản hồi
        data = response.json()

        # Kiểm tra xem phản hồi có chứa dữ liệu không
        if not data or len(data) < 2:
            print("Lỗi: Không nhận được dữ liệu từ API của Ngân hàng Thế giới.")
            return

        # Trích xuất dữ liệu thực tế từ phản hồi
        data = data[1]

        # Kiểm tra xem dữ liệu có trống không
        if not data:
            print("Lỗi: Không có mục dữ liệu trong phản hồi.")
            return

        # Chuyển đổi danh sách các từ điển thành DataFrame
        df = pd.json_normalize(data)

        # Chỉ giữ lại các cột quan trọng
        df = df[['country.value', 'date', 'value']]

        # Đổi tên cột để dễ hiểu hơn
        df.columns = ['Country', 'Year', 'GDP']

        # Lưu DataFrame vào tệp Excel
        df.to_excel(file_path, index=False)

        print(f"Dữ liệu đã được tải và lưu vào: {file_path}")

    except requests.exceptions.ConnectionError:
        print("Lỗi: Không thể kết nối với API của Ngân hàng Thế giới. Vui lòng kiểm tra kết nối internet của bạn và thử lại.")
    except requests.exceptions.HTTPError as e:
        print(f"Lỗi HTTP: {e}")
    except Exception as e:
        print(f"Có lỗi không mong muốn xảy ra: {e}")

def rearrange_and_save(file_path):
    # Tải dữ liệu từ Sheet1
    df = pd.read_excel(file_path, sheet_name='Sheet1')

    # Loại bỏ các mục trùng lặp
    df = df.drop_duplicates(subset=['Country', 'Year'])

    # Pivoting DataFrame để có các năm làm cột
    df_pivot = df.pivot(index='Country', columns='Year', values='GDP')

    # Đặt lại index và thêm cột 'Số thứ tự'
    df_pivot.reset_index(inplace=True)
    df_pivot.insert(0, 'Index', range(1, len(df_pivot) + 1))

    # Lưu dữ liệu đã sắp xếp vào một Sheet mới (Sheet2) trong tệp Excel
    new_file_path = 'downloaded_data\\World_Bank_Data.xlsx'
    with pd.ExcelWriter(new_file_path, engine='openpyxl', mode='a') as writer:
        # Thử loại bỏ Sheet nếu nó tồn tại
        try:
            writer.book.remove(writer.book['Sheet2'])
        except KeyError:
            pass

        # Lưu DataFrame vào tệp Excel
        df_pivot.to_excel(writer, sheet_name='Sheet2', index=False)

def clean_data_advanced(file_path):
    # Tải dữ liệu từ Sheet2
    df = pd.read_excel(file_path, sheet_name='Sheet2')

    # Các cột cần làm sạch (các cột GDP)
    columns_to_clean = df.columns[2:]

    # Tạo một bản sao của DataFrame để lưu dữ liệu đã làm sạch
    df_cleaned = df.copy()

    for column in columns_to_clean:

        df_cleaned[column].interpolate(method='linear', inplace=True)
        
        z_scores = np.abs(stats.zscore(df_cleaned[column]))

        outlier_threshold = 3

        df_cleaned[column] = np.where(z_scores >= outlier_threshold, np.nan, df_cleaned[column])

        percentile_95 = df_cleaned[column].quantile(0.95)
        df_cleaned[column] = np.where(df_cleaned[column] > percentile_95, percentile_95, df_cleaned[column])

    # Loại bỏ các dòng chứa giá trị NaN
    df_cleaned.dropna(inplace=True)

    # Đặt lại index
    df_cleaned.reset_index(drop=True, inplace=True)

    # Lưu dữ liệu đã làm sạch vào Sheet3 trong tệp Excel
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
        # Thử loại bỏ Sheet nếu nó tồn tại
        try:
            writer.book.remove(writer.book['Sheet3'])
        except KeyError:
            pass

        # Lưu DataFrame vào tệp Excel
        df_cleaned.to_excel(writer, sheet_name='Sheet3', index=False)

    print("Quá trình làm sạch dữ liệu đã hoàn thành và lưu vào Sheet3.")

def calculate_statistics(file_path):
    # Load data from Sheet3
    df = pd.read_excel(file_path, sheet_name='Sheet3')

    # Extract available years from DataFrame columns
    available_years = [col for col in df.columns if isinstance(col, int)]

    # Create a new DataFrame for storing results with hierarchical columns
    columns = pd.MultiIndex.from_product([available_years, ['Trung bình', 'Phương sai', 'Hiệp phương sai', '']], names=['Year', 'Statistic'])
    result_df = pd.DataFrame(index=df['Country'], columns=columns)

    # Calculate statistics for each year
    for year in available_years:
        # Access columns directly using integers
        year_data = df[year]
        result_df.loc[:, (year, 'Trung bình')] = year_data.mean(axis=0)
        result_df.loc[:, (year, 'Phương sai')] = year_data.var(axis=0)
        result_df.loc[:, (year, 'Hiệp phương sai')] = year_data.cov(df[year])

    # Write results to a new sheet (Sheet4) in the Excel file
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
        # Remove the sheet if it exists
        if 'Sheet4' in writer.book.sheetnames:
            writer.book.remove(writer.book['Sheet4'])

        # Write the DataFrame to the Excel file with a header row
        result_df.to_excel(writer, sheet_name='Sheet4', index=True, header=True)

    print("Statistics calculated and written to Sheet4.")

def analyze_and_visualize(file_path):
    # Load the data from Sheet3
    df = pd.read_excel(file_path, sheet_name='Sheet3')

    # Candlestick plot
    plt.figure(figsize=(14, 8))
    for index, row in df.iterrows():
        plt.plot(df.columns[2:], row[2:], label=row['Country'], marker='o')

    plt.title('Candlestick Plot of GDP Over Years for Each Country')
    plt.xlabel('Year')
    plt.ylabel('GDP')
    plt.legend()
    plt.show()

    # Print descriptive statistics
    print("Descriptive Statistics:")
    print(df.describe())

def random_forest_regression_all_countries(file_path, from_year_train, to_year_train, from_year_predict, to_year_predict):
    # Load full data from Excel file
    df_full = pd.read_excel(file_path, sheet_name='Sheet3')
    print(f"dataframe {df_full}")

    # Convert column names to strings
    df_full.columns = df_full.columns.astype(str)
    print("Column Names in df_full:", df_full.columns)

    # Extract feature columns for training
    feature_columns_train = [str(year) for year in range(from_year_train, to_year_train + 1)]

    # Initialize a DataFrame to store the predictions
    df_predictions = df_full.copy()

    # Check if the columns exist before assigning NaN values
    for year in range(from_year_predict, to_year_predict + 1):
        if str(year) not in df_predictions.columns:
            df_predictions[str(year)] = np.nan

    print("Updated Column Names in df_predictions:", df_predictions.columns)

    # Initialize a list to store residuals
    residuals = []
    errors = []
    
    # Iterate through each country for prediction
    for country in df_full['Country'].unique():
        # Iterate through each prediction year
        for year in range(from_year_predict, to_year_predict + 1):
            # Create a list of columns to be used for prediction
            pred_columns = [str(year - i) for i in range(1, len(feature_columns_train) + 1)]

            # Check if the columns exist in the DataFrame
            if all(col in df_full.columns for col in pred_columns):
                # Split data into training set
                X_train = df_full.loc[df_full['Country'] == country, feature_columns_train]
                y_train = df_full.loc[df_full['Country'] == country, str(year - 1)]  

                # Create and train the model
                model = RandomForestRegressor(n_estimators=100, random_state=42)
                model.fit(X_train, y_train)

                # Make predictions for the future year using the last available data
                X_pred = df_full.loc[df_full['Country'] == country, pred_columns].squeeze()

                # Check if there is at least one non-NaN value in X_pred
                if not X_pred.empty and not X_pred.isnull().all().all():
                    # Impute missing values using the mean (you can choose other strategies)
                    imputer = SimpleImputer(strategy='mean')
                    X_pred_imputed = imputer.fit_transform(X_pred.values.reshape(1, -1))  # Reshape to 2D array

                    # Make predictions for the future year using all the features used for training
                    future_prediction = model.predict(X_pred_imputed)

                    # Update the df_predictions DataFrame with the predicted value
                    df_predictions.loc[(df_predictions['Country'] == country), str(year)] = future_prediction

                    # Add the predicted value to the training data for the next iteration
                    df_full.loc[(df_full['Country'] == country), str(year)] = future_prediction

                    # Calculate the residual and add it to the list
                    residual = y_train - future_prediction
                    residuals.extend(residual)
                    
                    # Calculate error metrics
                    mae = mean_absolute_error(y_train, future_prediction)
                    mse = mean_squared_error(y_train, future_prediction)
                    rmse = sqrt(mse)
                    errors.append([country, year, mae, mse, rmse])

    # Combine training and prediction DataFrames
    df_full_combined = pd.concat([df_full, df_predictions.drop(columns=['Country'])], axis=1)

    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
        sheet_name = 'Sheet5'

        if sheet_name in writer.book.sheetnames:
            writer.book.remove(writer.book[sheet_name])

        df_predictions.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # Write error metrics to Sheet6
        df_errors = pd.DataFrame(errors, columns=['Country', 'Year', 'MAE', 'MSE', 'RMSE'])
        print("data frame errors: \n", df_errors)
        df_errors = df_errors.set_index(['Country', 'Year']).stack().reset_index()
        df_errors.columns = ['Country', 'Year', 'Statistic', 'Value']
        df_errors = df_errors.pivot_table(index='Country', columns=['Year', 'Statistic'], values='Value').reset_index()
        df_errors.columns = [' '.join(str(col_item) for col_item in col).strip() for col in df_errors.columns.values]
        df_errors.to_excel(writer, sheet_name='Sheet6', index=False)

    print("Final Predictions:")
    print(df_predictions)

    # Plot the residuals
    plt.hist(residuals, bins=20, edgecolor='black', density=True)
    plt.title('Histogram of Residuals')
    plt.xlabel('Residual')
    plt.ylabel('Frequency')
    plt.show()
    
    # Calculate histogram
    counts, bin_edges = np.histogram(residuals, bins=20, density=True)

    # Convert counts to frequencies
    frequencies = counts / sum(counts)

    # Plot histogram using bar chart
    plt.bar(bin_edges[:-1], frequencies, width=np.diff(bin_edges), edgecolor="black")
    plt.title('Histogram of Residuals')
    plt.xlabel('Residual')
    plt.ylabel('Frequency (%)')
    plt.show()

    return df_full_combined

def main():
    directory_path = 'downloaded_data'
    if not os.path.exists(directory_path):
        os.makedirs(directory_path)
        
    file_path = os.path.join(directory_path, 'World_Bank_Data.xlsx')

    from_year_train = int(input("Enter beginning year for training: "))
    to_year_train = int(input("Enter end year for training: "))
    from_year_predict = int(input("Enter year to predict: "))
    to_year_predict = int(input("Enter end year to predict: "))

    download_gdp_data('NY.GDP.MKTP.CD', from_year_train, to_year_train, file_path)

    rearrange_and_save(file_path)

    # Clean data
    clean_data_advanced(file_path)

    # Analyze data
    analyze_and_visualize(file_path)

    calculate_statistics(file_path)

    # Load data into DataFrame
    df = pd.read_excel(file_path, sheet_name='Sheet3')
    
    # Convert column names to strings when reading data from the file
    df.columns = df.columns.astype(str)

    # Print DataFrame information
    print("DataFrame Information:")
    
    print(df)
    print(df.columns)
    
    random_forest_regression_all_countries(file_path, from_year_train, to_year_train, from_year_predict, to_year_predict)
    
if __name__ == "__main__":
    main()


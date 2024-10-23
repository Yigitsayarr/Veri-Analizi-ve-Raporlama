import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Excel dosyasını oku
data = pd.read_excel(r'C:\Users\yigit\Desktop\prj\sales_data.xlsx')

# Verilerin ilk 5 satırını inceleyelim
print(data.head())

# Kategori bazında toplam satışlar
category_sales = data.groupby('Category')['Sales'].sum()

# Ortalama satışlar
average_sales = data['Sales'].mean()

print("Kategori Bazında Satışlar:")
print(category_sales)

print("\nOrtalama Satışlar:", average_sales)

# Satış verilerini çubuk grafik olarak görselleştir
plt.figure(figsize=(10,6))
sns.barplot(x=category_sales.index, y=category_sales.values)
plt.title('Kategori Bazında Toplam Satışlar')
plt.xlabel('Kategori')
plt.ylabel('Toplam Satış')
plt.tight_layout()
plt.savefig('category_sales_chart.png')  # Grafiği kaydediyoruz
plt.show()

# Yeni bir Excel dosyasına raporu yazma
with pd.ExcelWriter('sales_report.xlsx', engine='openpyxl') as writer:
    data.to_excel(writer, sheet_name='Raw Data', index=False)
    category_sales.to_excel(writer, sheet_name='Category Sales')
    
    # Excel dosyasına ortalama satışları yazdırma
    workbook  = writer.book
    worksheet = workbook.create_sheet('Summary')
    worksheet['A1'] = 'Ortalama Satışlar'
    worksheet['B1'] = average_sales

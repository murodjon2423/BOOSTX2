import pandas as pd
import matplotlib.pyplot as plt

def analyze_store_data(file_path, output_path):
    # Excel faylini yuklash
    df = pd.read_excel(file_path)
    
    # 1️⃣ Sotuvchi uchun tahlillar
    # Eng ko‘p sotilgan mahsulotlar (Miqdorga ko‘ra)
    top_products = df.groupby("Mahsulot nomi")["Miqdor"].sum().reset_index()
    top_products = top_products.sort_values(by="Miqdor", ascending=False)
    
    # Kunlik savdo miqdori
    daily_sales = df.groupby("Sana")["Jami summa (so‘m)"].sum().reset_index()
    
    # To‘lov turlari taqsimoti
    payment_distribution = df["To‘lov turi"].value_counts().reset_index()
    payment_distribution.columns = ["To‘lov turi", "Miqdor"]
    
    # 2️⃣ Hisobchi uchun tahlillar
    # Umumiy tushum
    total_revenue = df["Jami summa (so‘m)"].sum()
    
    # O'rtacha xarid summasi
    average_purchase = df["Jami summa (so‘m)"].mean()
    
    # Eng foydali mahsulotlar (Jami summaga ko‘ra)
    top_revenue_products = df.groupby("Mahsulot nomi")["Jami summa (so‘m)"].sum().reset_index()
    top_revenue_products = top_revenue_products.sort_values(by="Jami summa (so‘m)", ascending=False)
    
    # Diagrammalar yaratish
    plt.figure(figsize=(10, 5))
    plt.bar(top_products["Mahsulot nomi"].head(10), top_products["Miqdor"].head(10), color='skyblue')
    plt.xticks(rotation=45)
    plt.xlabel("Mahsulot nomi")
    plt.ylabel("Sotilgan miqdor")
    plt.title("Eng ko‘p sotilgan mahsulotlar")
    plt.savefig("top_products.png")
    plt.close()
    
    plt.figure(figsize=(10, 5))
    plt.bar(daily_sales["Sana"], daily_sales["Jami summa (so‘m)"], color='orange')
    plt.xticks(rotation=45)
    plt.xlabel("Sana")
    plt.ylabel("Kunlik tushum")
    plt.title("Kunlik savdo miqdori")
    plt.savefig("daily_sales.png")
    plt.close()
    
    plt.figure(figsize=(6, 6))
    plt.pie(payment_distribution["Miqdor"], labels=payment_distribution["To‘lov turi"], autopct='%1.1f%%', colors=['lightblue', 'lightcoral', 'lightgreen'])
    plt.title("To‘lov turlari taqsimoti")
    plt.savefig("payment_distribution.png")
    plt.close()
    
    # Hisobchi uchun tushum statistikasi
    revenue_stats = pd.DataFrame({
        "Ko‘rsatkichlar": ["Umumiy tushum", "O‘rtacha xarid summasi"],
        "Qiymatlar": [total_revenue, average_purchase]
    })
    
    plt.figure(figsize=(6, 4))
    plt.bar(revenue_stats["Ko‘rsatkichlar"], revenue_stats["Qiymatlar"], color='purple')
    plt.xlabel("Ko‘rsatkichlar")
    plt.ylabel("Qiymat (so‘m)")
    plt.title("Hisobchi uchun asosiy statistikalar")
    plt.savefig("revenue_stats.png")
    plt.close()
    
    # Yangi Excel faylni yaratish
    with pd.ExcelWriter(output_path) as writer:
        df.to_excel(writer, sheet_name="Asosiy Ma'lumotlar", index=False)
        top_products.to_excel(writer, sheet_name="Ko‘p Sotilgan Mahsulotlar", index=False)
        daily_sales.to_excel(writer, sheet_name="Kunlik Savdo", index=False)
        payment_distribution.to_excel(writer, sheet_name="To‘lov Turlari", index=False)
        top_revenue_products.to_excel(writer, sheet_name="Foydali Mahsulotlar", index=False)
        revenue_stats.to_excel(writer, sheet_name="Hisobchi Statistikasi", index=False)
    
    print(f"Tahlil yakunlandi! Natijalar {output_path} fayliga saqlandi.")
    print("Diagrammalar saqlandi: top_products.png, daily_sales.png, payment_distribution.png, revenue_stats.png")

# Dastur ishga tushirilganda ishlaydi
def main():
    input_file = "do'kon_oziq_ovqat.xlsx"  # Yuklab olingan fayl nomi
    output_file = "do'kon_tahlil.xlsx"  # Tahlil natijalari saqlanadigan fayl
    analyze_store_data(input_file, output_file)

if __name__ == "__main__":
    main()

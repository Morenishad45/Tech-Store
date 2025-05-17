import pandas as pd
import datetime
import os

excel_file_path = r"C:\Users\moren\OneDrive\psp\pspexcel2.xlsx"
fallback_excel_path = os.path.join(os.path.expanduser("~"), "Desktop", "pspexcel2.xlsx")

def load_products():
    try:
        df = pd.read_excel(excel_file_path, sheet_name='Products')
        return df
    except Exception as e:
        print(f"Error loading products: {e}")
        return pd.DataFrame(columns=['Company', 'Model', 'Sub-Model', 'Variant', 'Price', 'Quantity'])

def load_products():
    try:
        df = pd.read_excel(excel_file_path, sheet_name='Products')
        df.columns = df.columns.str.strip() 
        return df
    except Exception as e:
        print(f"Error loading products: {e}")
        return pd.DataFrame(columns=['Company', 'Model', 'Sub-Model', 'Variant', 'Price', 'Quantity'])

def load_transactions():
    try:
        df = pd.read_excel(excel_file_path, sheet_name='Transactions')
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        print(f"Error loading transactions: {e}")
        return pd.DataFrame(columns=[
            'Date', 'Company', 'Model', 'Sub-Model',
            'Variant', 'Quantity', 'Price per Unit', 'Total Price', 'Type'
        ])

def save_products(df):
    try:
        with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name='Products', index=False)
    except Exception as e:
        print(f"Error saving products: {e}")

def save_transactions(df):
    try:
        with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name='Transactions', index=False)
    except Exception as e:
        print(f"Error saving transactions: {e}")

def save_to_excel(data):
    try:
        df = load_transactions() 
        df = pd.concat([df, pd.DataFrame(data)], ignore_index=True)
        save_transactions(df)
        print(f" Transaction saved")
    except Exception as e:
        print(f"Failed to write to primary Excel path: {e}")
        try:
            df = pd.DataFrame(data)
            df.to_excel(fallback_excel_path, index=False)
            print(f"Saved to fallback path instead: {fallback_excel_path}")
        except Exception as e2:
            print(f"Failed to write to fallback path too: {e2}")


def view_available_products():
    df = load_products()
    if df.empty:
        print("\nNo products available.")
        return

    for col in df.select_dtypes(include='object').columns:
        df[col] = df[col].str.strip()

    print("\n" + "="*40)
    print("         Available Products")
    print("="*40)

    grouped = df.groupby(['Company', 'Model', 'Sub-Model'])

    for (company, model, sub_model), group in grouped:
        print(f"\nCompany: {company}")
        print(f"  Model: {model}")
        print(f"    Sub-Model: {sub_model}")
        for _, row in group.iterrows():
            print(f"      - Variant: {row['Variant']:<20} | Price: ₹{row['Price']:<8} | Quantity: {row['Quantity']}")
        print("-" * 40)


def replace_product():
    transactions_df = load_transactions()

    if transactions_df.empty:
        print("No transactions found.")
        return

    transactions_df.dropna(subset=['Company', 'Model', 'Sub-Model', 'Variant'], inplace=True)

    print("\nTransaction History (Choose a product to mark as replaced):")
    print("=" * 90)
    for idx, row in transactions_df.iterrows():
        print(f"{idx + 1}. {row['Date']} | {row['Company']} - {row['Model']} - {row['Sub-Model']} - {row['Variant']} | Qty: {row['Quantity']} | Type: {row['Type']}")
    print("=" * 90)

    try:
        choice = int(input("Enter the number of the product to replace: ")) - 1
        if not (0 <= choice < len(transactions_df)):
            print("Invalid selection.")
            return
    except ValueError:
        print(" Please enter a valid number.")
        return

    selected = transactions_df.iloc[choice]
    company, model, sub_model, variant = selected['Company'], selected['Model'], selected['Sub-Model'], selected['Variant']

    df_products = load_products()

    if variant not in df_products['Variant'].values:
        print(f"Variant '{variant}' not found in inventory.")
        return

    df_products = df_products[df_products['Variant'] != variant]
    save_products(df_products)
    print(f"Product '{company} {model} {sub_model} {variant}' has been replaced and removed from inventory.")

    transaction_data = {
        'Date': datetime.datetime.now().strftime("%Y-%m-%d"),
        'Company': company,
        'Model': model,
        'Sub-Model': sub_model,
        'Variant': variant,
        'Quantity': 0,
        'Price per Unit': 0,
        'Total Price': 0,
        'Type': 'Replaced'
    }

    save_to_excel([transaction_data])
    print("Replacement transaction logged.")

def purchase_flow():
    def choose_option(options, label):
        print(f"\nSelect {label}:")
        for i, option in enumerate(options, 1):
            print(f"{i}. {option}")
        while True:
            try:
                choice = int(input(f"Enter {label} number: "))
                if 1 <= choice <= len(options):
                    return options[choice - 1]
                else:
                    print("Invalid number.")
            except ValueError:
                print("Enter a valid number.")

    df = load_products()
    products = {}
    for _, row in df.iterrows():
        company = row['Company']
        model = row['Model']
        sub_model = row['Sub-Model']
        variant = row['Variant']
        price = row['Price']
        quantity = row['Quantity']

        products.setdefault(company, {}).setdefault(model, {}).setdefault(sub_model, {})[variant] = {
            'price': price,
            'quantity': quantity
        }

    cart = []

    while True:
        company = choose_option(list(products.keys()), "Company")
        model = choose_option(list(products[company].keys()), "Model")
        sub_model = choose_option(list(products[company][model].keys()), "Sub-Model")
        variant = choose_option(list(products[company][model][sub_model].keys()), "Variant")

        available_qty = products[company][model][sub_model][variant]['quantity']
        print(f"Available Quantity: {available_qty}")

        while True:
            try:
                qty = int(input("Enter quantity to purchase: "))
                if 1 <= qty <= available_qty:
                    break
                else:
                    print("Enter quantity within available stock.")
            except ValueError:
                print("Enter a valid number.")

        price = products[company][model][sub_model][variant]['price']
        cart.append({
            'Company': company,
            'Model': model,
            'Sub-Model': sub_model,
            'Variant': variant,
            'Quantity': qty,
            'Price per Unit': price,
            'Total Price': price * qty
        })
        print(f"Added {qty} x {variant} to cart.")

        another = input("Add another product? (yes/no): ").lower()
        if another != 'yes':
            break

    if not cart:
        print("Cart is empty.")
        return

    confirm = input("Proceed to checkout? (yes/no): ").lower()
    if confirm == 'yes':
        total = sum(item['Total Price'] for item in cart)
        checkout_multiple(cart, total)

        print("\nCart Summary:")
        print("="*60)
        for item in cart:
            print(f"{item['Company']} {item['Model']} {item['Sub-Model']} {item['Variant']} - Qty: {item['Quantity']} - ₹{item['Total Price']:,.2f}")
        print("="*60)
        print(f"Total Paid: ₹{total:,.2f}")

def checkout(total_price, company, model, sub_model, variant, qty):
    transaction_data = {
        'Date': datetime.datetime.now().strftime("%Y-%m-%d"),
        'Company': company,
        'Model': model,
        'Sub-Model': sub_model,
        'Variant': variant,
        'Quantity': qty,
        'Price per Unit': total_price / qty,
        'Total Price': total_price,
        'Type': 'Purchase'
    }
    save_to_excel([transaction_data])
    print(f"Transaction successful. Total amount: ₹{total_price:,.2f}")
    df = load_products()
    mask = (
        (df['Company'] == company) &
        (df['Model'] == model) &
        (df['Sub-Model'] == sub_model) &
        (df['Variant'] == variant)
    )

    if not mask.any():
        print("Error: Product not found in inventory.")
        return
    df.loc[mask, 'Quantity'] = df.loc[mask, 'Quantity'] - qty
    df['Quantity'] = df['Quantity'].apply(lambda x: max(x, 0))
    save_products(df)
    updated_row = df[mask]
    if not updated_row.empty:
        print("\nUpdated Inventory for This Product:")
        print(updated_row.to_string(index=False))
    else:
        print("Product missing after update — check Variant uniqueness.")

def checkout_multiple(cart, total_price):
    transactions = []
    for item in cart:
        transactions.append({
            'Date': datetime.datetime.now().strftime("%Y-%m-%d"),
            'Company': item['Company'],
            'Model': item['Model'],
            'Sub-Model': item['Sub-Model'],
            'Variant': item['Variant'],
            'Quantity': item['Quantity'],
            'Price per Unit': item['Price per Unit'],
            'Total Price': item['Total Price'],
            'Type': 'Purchase'
        })

    save_to_excel(transactions)
    print("All transactions saved successfully.")

    df = load_products()
    for item in cart:
        mask = (
            (df['Company'] == item['Company']) &
            (df['Model'] == item['Model']) &
            (df['Sub-Model'] == item['Sub-Model']) &
            (df['Variant'] == item['Variant'])
        )
        df.loc[mask, 'Quantity'] -= item['Quantity']
    save_products(df)
    print("Stock updated for all items.")
    
def add_product():
    df = load_products()
    company = input("Enter Company: ")
    model = input("Enter Model: ")
    sub_model = input("Enter Sub-Model: ")
    variant = input("Enter Variant: ")
    try:
        price = float(input("Enter Price: "))
        quantity = int(input("Enter Quantity: "))
    except ValueError:
        print("Invalid price or quantity.")
        return
    new_product = pd.DataFrame([{
        'Company': company,
        'Model': model,
        'Sub-Model': sub_model,
        'Variant': variant,
        'Price': price,
        'Quantity': quantity
    }])
    df = pd.concat([df, new_product], ignore_index=True)
    save_products(df)
    print("Product added successfully.")
    
def update_quantity():
    df = load_products()
    
    if df.empty:
        print("No products found in inventory.")
        return

    print("\nAvailable Products to Update Quantity:")
    print("=" * 60)
    for idx, row in df.iterrows():
        print(f"{idx + 1}. {row['Company']} > {row['Model']} > {row['Sub-Model']} > {row['Variant']} | Price: ₹{row['Price']} | Quantity: {row['Quantity']}")
    print("=" * 60)

    try:
        choice = int(input("Enter the number of the product to update quantity: "))
        if 1 <= choice <= len(df):
            selected_idx = choice - 1
        else:
            print("Invalid choice number.")
            return
    except ValueError:
        print("Please enter a valid number.")
        return

    try:
        additional_qty = int(input("Enter quantity to add: "))
        if additional_qty < 0:
            print("Cannot add negative quantity.")
            return
    except ValueError:
        print("Invalid quantity input.")
        return

    df.at[selected_idx, 'Quantity'] += additional_qty
    save_products(df)

    print(f"Quantity updated. New stock for {df.at[selected_idx, 'Variant']}: {df.at[selected_idx, 'Quantity']}")

def delete_variant():
    df = load_products()
    variant = input("Enter Variant to delete: ")
    if variant not in df['Variant'].values:
        print("Variant not found.")
        return
    df = df[df['Variant'] != variant]
    save_products(df)
    print("Variant deleted successfully.")

def monthly_sales_report():
    try:
        df = load_transactions()
    except Exception as e:
        print(f"Error loading transactions: {e}")
        return

    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    current_month = datetime.datetime.now().strftime("%Y-%m")
    df = df[df['Date'].dt.strftime('%Y-%m') == current_month]

    if df.empty:
        print("No transactions found for this month.")
        return

    purchase_df = df[df['Type'] == 'Purchase']

    if purchase_df.empty:
        print("No valid purchases for this month.")
        return

    summary = purchase_df.groupby(['Company', 'Model', 'Sub-Model', 'Variant']).agg({
        'Quantity': 'sum',
        'Total Price': 'sum'
    }).reset_index()

    print("\nMonthly Sales Summary (Excludes Replacements Only):")
    print("=" * 80)
    print(f"{'Company':<20}{'Model':<20}{'Sub-Model':<20}{'Variant':<20}{'Quantity':<10}{'Total Price'}")
    print("=" * 80)
    for _, row in summary.iterrows():
        print(f"{row['Company']:<20}{row['Model']:<20}{row['Sub-Model']:<20}{row['Variant']:<20}{row['Quantity']:<10}{row['Total Price']}")
    print("=" * 80)

    total_sales = summary['Total Price'].sum()
    print(f"\nTotal Sales for this month: ₹{total_sales:,.2f}")

def seller():
    print("\nSeller Menu:")
    print("1. Add Product")
    print("2. Update Quantity")
    print("3. Delete Product Variant")
    print("4. Monthly Sales Report")
    
    choice1 = input("Enter Choice: ")
    
    if choice1 == '1':
        add_product()
    elif choice1 == '2':
        update_quantity()
    elif choice1 == '3':
        delete_variant()
    elif choice1 == '4':
        monthly_sales_report()
    else:
        print("Invalid choice.")

def main():
    while True:
        print("\nMain Menu:")
        print("1. View Products")
        print("2. Purchase Product")
        print("3. Replace Product")
        print("4. Seller Window")
        print("5. Exit")
        
        choice = input("Enter choice: ")
        
        if choice == '1':
            view_available_products()
        elif choice == '2':
            purchase_flow()
        elif choice == '3':
            replace_product()
        elif choice == '4':
            a=input("Enter Password-")
            if a == 'techstore.com':
                seller()
            else :
                main()
        elif choice == '5' :
            print("Thanks for Visiting")
            break
        else :
            print("Invalid choice")

main()
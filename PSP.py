import pandas as pd
import pickle

monthly_income = {}
monthly_budget = {}
records = []
current_balance = {}


def get_number(msg):
    while True:
        try:
            return int(input(msg))
        except:
            print("Enter valid number")

def save_data():
    data = {
        "monthly_income": monthly_income,
        "monthly_budget": monthly_budget,
        "records": records,
        "current_balance": current_balance
    }

    with open("finance_data.pkl", "wb") as file:
        pickle.dump(data, file)

def load_data():
    global monthly_income, monthly_budget, records, current_balance

    try:
        with open("finance_data.pkl", "rb") as file:
            data = pickle.load(file)

            monthly_income = data["monthly_income"]
            monthly_budget = data["monthly_budget"]
            records = data["records"]
            current_balance = data["current_balance"]

        print("Previous data loaded successfully  ")

    except FileNotFoundError:
        print("No previous data found. Starting fresh ")

def add_income(month):
    if month in monthly_income:
        print("Income already added for this month ")
        return monthly_income[month]

    income = {}
    n = get_number("How many income sources? ")

    for i in range(n):
        name = input("Enter source name: ")
        amt = get_number("Enter income amount: ")
        income[name] = amt

    monthly_income[month] = income
    current_balance[month] = sum(income.values())

    save_data()

    return income

def add_budget(month):
    if month in monthly_budget:
        print("Budget already added for this month!")
        return monthly_budget[month]

    budget = {}
    n = get_number("How many budget categories? ")

    for i in range(n):
        cat = input("Enter budget category: ")
        amt = get_number("Enter budget amount: ")
        budget[cat] = amt

    monthly_budget[month] = budget

    total_budget = sum(budget.values())
    total_income = sum(monthly_income.get(month, {}).values())

    print("\nBUDGET STATUS")

    if total_budget > total_income:
        print("Warning: Budget is greater than Income ")
        print("Income:", total_income)
        print("Budget:", total_budget)
        print("Exceeded by:", total_budget - total_income)

    elif total_budget == total_income:
        print("Budget equals Income (No Saving Possible)")

    else:
        print("Budget is within Income")
        print("Possible Saving:", total_income - total_budget)

    save_data()

    return budget

def add_expense(month):

    if month not in monthly_budget:
        print("Please set income and budget first!")
        return

    date =  input("Enter date (DD-MM-YYYY): ")

    budget = monthly_budget[month]

    print("\n1. Budget Category Expense")
    print("2. Other Expense")

    choice = input("Enter choice: ")

    if choice == "1":

        print("Available Categories:", list(budget.keys()))

        category = input("Enter category: ")

        if category not in budget:
            print("Invalid category!")
            return

        amount = get_number("Enter amount: ")

        records.append((month, date, category, amount))

        current_balance[month] -= amount

        print("Remaining Balance:", current_balance[month])

        if current_balance[month] < 0:
            print("You are in LOSS!")

        total_spent = 0

        for r in records:
            if r[0] == month and r[2] == category:
                total_spent += r[3]

        limit = budget[category]

        if total_spent > limit:
            print("Budget exceeded by:", total_spent - limit)
        else:
            print("Remaining Budget:", limit - total_spent)

        save_data()

    elif choice == "2":

        n = get_number("How many other expenses? ")

        for i in range(n):

            category = input("Enter category: ")
            amount = get_number("Enter amount: ")

            records.append((month, date, category, amount))

            current_balance[month] -= amount

            print("Remaining Balance:", current_balance[month])

            if current_balance[month] < 0:
                print("You are in LOSS!")

        save_data()

    else:
        print("Invalid choice!")

def create_report():

    if not monthly_income:
        print("No income data available!")
        return
    if records:
        df = pd.DataFrame(
            records,
            columns=["Month", "Date", "Category", "Expense"]
        )

        df = df.sort_values(by="Month")
    else:
        df = pd.DataFrame(
            columns=["Month", "Date", "Category", "Expense"]
        )

    summary_list = []

    months = sorted(monthly_income.keys(), key=int)

    prev_balance = 0

    for month in months:

        month_data = (
            df[df["Month"] == month]
            if not df.empty
            else pd.DataFrame()
        )

        total_income = sum(
            monthly_income.get(month, {}).values()
        )

        total_budget = sum(
            monthly_budget.get(month, {}).values()
        )

        if not month_data.empty:
            total_expense = month_data["Expense"].sum()
        else:
            total_expense = 0

        saving = total_income - total_expense

        if saving < 0:
            status = "Loss"
            saving_display = abs(saving)
        else:
            status = "Saving"
            saving_display = saving

        remaining = prev_balance + saving
        prev_balance = remaining

        summary_list.append({
            "Month": month,
            "Total Income": total_income,
            "Total Budget": total_budget,
            "Total Expense": total_expense,
            "Saving/Loss": saving_display,
            "Status": status,
            "Remaining Balance": remaining
        })

    df_summary = pd.DataFrame(summary_list)

    file_name = "finance_tracker_final.xlsx"

    with pd.ExcelWriter(file_name) as writer:

        df.to_excel(
            writer,
            sheet_name="Expenses",
            index=False
        )

        df_summary.to_excel(
            writer,
            sheet_name="Summary",
            index=False
        )

    print("\nExcel File Saved Successfully!")
    print("File Name:", file_name)

load_data()

while True:

    print("\nFINANCE TRACKER MENU ")
    print("1. Add Income & Budget")
    print("2. Add Expense")
    print("3. Create Report")
    print("4. Exit")

    ch = input("Enter choice: ")

    if ch == "1":

        month = input("Enter month (MM): ")

        add_income(month)

        add_budget(month)

    elif ch == "2":

        month = input("Enter month (MM): ")

        add_expense(month)

    elif ch == "3":

        create_report()

    elif ch == "4":

        create_report()

        save_data()

        print("\nData saved successfully!")
        print("Program exited!")

        break

    else:
        print("Invalid choice!")
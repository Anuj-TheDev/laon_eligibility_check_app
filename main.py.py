import os
from rich.console import Console
from openpyxl import Workbook,load_workbook
from rich.progress import track
from rich.panel import Panel
import sys
import time
console=Console()
file_name="bank_entry.xlsx"


# --------------------progress bar-------------------
def progress_bar(val=1):
    for _ in track(range(val), description="Processing...."):
        time.sleep(1)


# -------------------banner-------------------------------
def banner():
    print("""████████████████████████████████████████████████████████
█                 ■ AS FINSERV ■                       █
█               LOANCHECK SYSTEM ⚡                    █
████████████████████████████████████████████████████████
""")

# --------------------typewriter-----------------------------   
def typewriter(text, delay=0.05):
    for char in text:
        sys.stdout.write(char)
        sys.stdout.flush()  # Forces the terminal to show the character immediately
        time.sleep(delay)
    print()


# ----------------------Panel-------------------------------
def panel():
    with console.status("[bold cyan]Running the eligibility check....[/bold cyan]",spinner="dots"):
        time.sleep(2)




# -----------------------cibl score----------------------------
def cibl_score():
    while True:
        try:
            cibl = int(console.input("[bold yellow]>CIBIL Score : [/bold yellow]"))
            
            if cibl < 400 or cibl > 900:
                console.print("[bold red]Error: CIBIL score must be between 400 and 900![/bold red]")
            else:
                return cibl
                

        except ValueError:
            console.print("[bold red]Enter digits, try again![/bold red]")



banner()
typewriter("Can You Get a Loan? Find Out Now.....",delay=0.05)   
print()
print()
print()
console.print("─" * 65 + "[bold blue] NEW APPLICATION [/bold blue]" + "─" * 65, style="blue")


# --------------------------excel file opening-------------------------------------
if not os.path.exists(file_name):
    wb=Workbook()
    sheet=wb.active
    sheet.title="sheet"
    headers=["Name","Age","Type","Monthly Salary","CIBL SCORE","Existing Loans","loan Amount","Tenure","Credit Band","Rate","EMI","FOIR","Status"]
    sheet.append(headers)
    wb.save(file_name)
    wb.close()


# ----------------------------user data-------------------------------
while True:
    name=console.input("[bold yellow]>Enter Full Name :[/bold yellow]")
    if name.replace(" ","").isalpha():
        break
    else:
        console.print("[bold red]invalid input try again!![/bold red]")
while True:
        try:
            age=int(console.input("[bold yellow]>Enter current age :[/bold yellow]"))
            if age<=18:
                console.print(f"[bold red]✘Sorry {name} you cannot apply for loan!![/bold red]")
                sys.exit()
            elif age>=60:
                console.print(f"[bold red]✘Sorry {name} you cannot apply for loan!![/bold red]")
                sys.exit()
            else:
                break
        except ValueError:
            console.print(f"[bold red]✘Invalid input {name}[/bold red]")        
  

type=["salaried","student","self employed"]
while True:
    employed_type=console.input("[bold Blue]>Enter Type(student,salaried,self employed) :[/bold blue]")
    if employed_type not in type:
        console.print(f"[bold red]✘Enter valid input!![/bold red]")
    else:
        break    

while True:    
    try:
            
            monthly_income=int(console.input("[bold yellow]>Enter Monthly income :[/bold yellow]"))
            

            if monthly_income<=0:
                console.print(f"[bold red]✘You are not eligible!![/bold red]")
                break
            cibl=cibl_score()
            break
            
                
    except ValueError:
        console.print(f"[bold red]Invalid Number!![/bold red]")    
while True:
    try:
        existing_loan=int(console.input("[bold yellow]>Kitne Existing loan chal rahe hai ? :[/bold yellow]"))
        if existing_loan < 0:
            console.print("[bold red]✘Invalid input!no of loans cannot be negative....[/bold red]")
        break
    except ValueError:
        console.print("[bold red]✘Invalid amount[/bold red]")


while True:
    try:
        loan_amount=int(console.input("[bold yellow]>Loan amount required($) :[/bold yellow]"))
        if loan_amount<=0:
            console.print("[bold red]✘Loan amount cannot be zero or negative!![/bold red]")
        break    
    except ValueError:
        console.print("[bold red]✘Inavlid input enter numbers only!![/bold red]")


while True:
    try:
        months=[12,24,36,48,60]
        tenure=int(console.input("[bold yellow]>Tenure in months (12/24/36/48/60) :[/bold yellow]"))
        if tenure not in months:
            console.print("[bold red]✘Invalid choice ,choose from (12/24/36/48/60).[/bold red]")
        break
    except ValueError:
        console.print("[bold red]✘Invalid input enter only numbers.[/bold red]")    



panel()
progress_bar(2)

# ----------------------------------------------screening result------------------------------------------------------------




# ---------------checking the cibil score on that basis giveing the credit and interest rate--------------------
if 750<= cibl <=900:
    credit_band="Excellent"
    interest_rate=11.0
elif 700<=cibl <=750:
    credit_band="Good"
    interest_rate=13.0
elif 650<= cibl <=750:
    credit_band="Fair"
    interest_rate=16.0
else:
    credit_band="Poor"
    interest_rate=0.0            







# ---------------------------------------calculating the Emi and FOIR Final status----------------------------------

rate=(interest_rate/12)/100

if rate>0:
    monthly_emi= (loan_amount*rate* ((1+rate)**tenure))/(((1+rate)**tenure)-1)    #  ** INDICATES THE POWER TO THE EXPONENT OPERATORI 
else:
    monthly_emi=loan_amount/tenure

existing_emi=existing_loan*5000

foir=((existing_emi + monthly_emi) / monthly_income)* 100




if credit_band=="Poor" or foir >50 or employed_type=="student":
    status="REJECTED"
else:
    status="APPROVED"  

# --------------------------generating id -------------------------------

plain_name=name.replace(" ","")
loan_id=f"ASB-LN-{plain_name[:3].upper()}{age}"

# -------------------------stacking all details into the dictionary-----------------------
# remains new and clean for every new customer


cust_details={
    "id": loan_id,
    "name":name.lower(),
    "band":credit_band,
    "rate":f"{interest_rate}% p.a.",
    "amount":f"{loan_amount:,.2f}",                       #FROMATTING NUMBER WITH THE COMMAS ASWELL ---:,.2f
    "tenure":f"{tenure} months",
    "emi":f"{monthly_income:,.2f}",
    "existing": f"{existing_emi:,.2f}",
    "foir":f"{foir:.2f}%",
    "status":status
}
if cust_details["status"]=="APPROVED":
    status_val="APPROVED"
else:
    status_val="REJECTED"    

# -----------------------displaying the result------------------------

report_body = (
    f"[bold white]{'Application ID':<16}[/bold white] :  [bold cyan]{cust_details['id']}[/bold cyan]\n"
    f"[bold white]{'Applicant':<16}[/bold white] :  {cust_details['name']}\n"
    f"[bold white]{'Credit Band':<16}[/bold white] :  {cust_details['band']}\n"
    f"[bold white]{'Interest Rate':<16}[/bold white] :  {cust_details['rate']}\n\n"
    f"[bold white]{'Loan Amount':<16}[/bold white] :  {cust_details['amount']}\n"
    f"[bold white]{'Tenure':<16}[/bold white] :  {cust_details['tenure']}\n"
    f"[bold white]{'Monthly EMI':<16}[/bold white] :  {cust_details['emi']}\n"
    f"[bold white]{'Existing EMIs':<16}[/bold white] :  {cust_details['existing']}\n"
    f"[bold white]{'FOIR':<16}[/bold white] :  {cust_details['foir']}\n\n"
    f"[bold white]{'STATUS':<16}[/bold white] :  {status_val}"
    )



panel = Panel(
    report_body,
    title="🏛️ LOANCHECK - ELIGIBILITY REPORT",
    title_align="left",
    expand=False,
    padding=(1, 4)
)

console.print(panel)


# ------------------------------------printing the last final message---------------------------------------
if cust_details["status"]=="APPROVED":
    console.print(f"\n[bold green]Congratulatins {cust_details['name']}! Apka Loan approve hogaya...[/bold green]")
else:
    console.print(f"\n[bold red]Sorry {cust_details['name']} Apka loan reject hogaya...!!!![/bold red]")


# ---------------------saving all the data in excel sheet-------------------

wb=load_workbook(file_name)
sheet=wb.active
sheet.append([name,
              age,
              employed_type,
              monthly_income,
              cibl,
              existing_loan,
              loan_amount,
              tenure,
              cust_details["band"],
              cust_details["rate"],
              cust_details["emi"],
              cust_details["foir"],
              cust_details["status"]
              ])
wb.save(file_name)
wb.close()













        
        



    


        

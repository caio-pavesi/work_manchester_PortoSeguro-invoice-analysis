# --- Imports --- #

import re
import os
import sys
import pandas
import datetime

from pathlib import Path

# Modules data
from modules.pdf import Pdf
from modules. log import Log
from modules.outlook import Outlook
from modules.metabase import Metabase

# Config data
from config.secrets import credentials
from config.automations import config


# --- Functions --- #

def get_extracts(account: str, search_folder: str, output_folder: str, search_date: str, target_file: str):
    
    outlook_app = Outlook(account)
    Outlook.output_folder = output_folder
    
    emails = outlook_app.search_emails(look_folder = search_folder, date_interval = search_date)
    if emails == []:
        Log().log(f"{datetime.datetime.now()} - 700000 [PRD] - No emails on the period\n")
        Log().log(f"{datetime.datetime.now()} - 700000 [PRD] - Process completed successfully\n")
        sys.exit(0)
    
    extracts = outlook_app.download_attachments(emails, target_file)
    Log().log(f"{datetime.datetime.now()} - 700000 [PRD] - Emails obtained\n")
    
    return extracts


def read_extracts(files: list, folder_path: Path):
    
    data = list()
    for file_name in files:    
        employees = list()
        
        file_path = folder_path / file_name
        content = Pdf(file_path).read_text().split('\n')
        
        for line in content:
            if re.findall(r'^Cobrança[.]+\s:\s[0-9.]+\s', line):
                company = re.split(r'^Cobrança[.]+\s:\s[0-9.]+\s', line)[1]
                company = company.split(" ")[0:2]
                company = " ".join(company).strip()
            elif re.findall(r"^\d+\s[A-Z'\s]+[0-9.,]+\s[0-9.,]+\s[0-9.,]+\s[0-9.,]+\s[0-9.,]+\s[0-9.,]+$", line):
                employee = re.findall(r"[A-Z]+", line)
                employee = " ".join(employee).strip()
                employees.append(employee)
        
        for employee in employees:
            data.append(
                {"Empresa": company
                ,"Colaborador": employee
                ,"Arquivo": file_path}
            )
    
    data = pandas.DataFrame(data)
    Log().log(f"{datetime.datetime.now()} - 700000 [PRD] - Insurance extract read\n")
    
    return data


def get_employees(username, password, domain, database, table):
    """Obtains employee data from metabase

    Args:
        username (str): login credential
        password (str): login credential
        domain (str): server domain
        database (int): database id
        table (int): table id

    Returns:
        data (pandas.DataFrame): Employee data
    """
    
    Metabase.username = username
    Metabase.password = password
    Metabase.domain = domain
    
    data = Metabase().get_table(database, table)
    data = pandas.DataFrame(data)
    data = data[["Nome Funcionario", "Data Demissao"]]
    Log().log(f"{datetime.datetime.now()} - 700000 [PRD] - Employee data obtained\n")
    
    return data


def do_analysis(extract_data: pandas.DataFrame, source_data: pandas.DataFrame):
    
    dismissed_employees = source_data.loc[~source_data["Data Demissao"].isna()].reset_index(drop = True)
    employees_insured = extract_data
    
    # Data cleaning (Remove accents)
    dismissed_employees.loc[:, ("Nome Funcionario")] = dismissed_employees.loc[:, ("Nome Funcionario")].str.upper()
    dismissed_employees.loc[:, ("Nome Funcionario")] = dismissed_employees.loc[:, ("Nome Funcionario")].str.normalize('NFKD')
    dismissed_employees.loc[:, ("Nome Funcionario")] = dismissed_employees.loc[:, ("Nome Funcionario")].str.encode('ascii', errors = 'ignore')
    dismissed_employees.loc[:, ("Nome Funcionario")] = dismissed_employees.loc[:, ("Nome Funcionario")].str.decode('utf-8')
    
    # DataFrame comparison
    data = dismissed_employees.merge(employees_insured, left_on = "Nome Funcionario", right_on = "Colaborador")
    data = data.loc[:, ("Empresa", "Nome Funcionario", "Data Demissao", "Arquivo")]
    Log().log(f"{datetime.datetime.now()} - 700000 [PRD] - Analysis done\n")
    
    return data


def send_analysis(account: str, recipient: str, analysis: pandas.DataFrame, folder: str):

    outlook_app = Outlook(account)    
    companies = analysis["Empresa"].unique()
    companies_html_body = list()
    
    files = os.listdir(folder)
    extracts = list()
    for file in files:
        file_path = folder / file
        extracts.append(file_path)
    
    for company in companies:
        company_data = analysis.loc[analysis["Empresa"] == company]
        company_analysis = company_data.loc[:, ("Nome Funcionario", "Data Demissao")].to_html()
        
        comapny_html_div = f"""
            <div>
                For the company {company} the following billed employees are terminated:
                <br/><br/>
                {company_analysis}
                <br/>
            <div/>
        """
        companies_html_body.append(comapny_html_div)
    
    companies = ", ".join(companies).strip()
    
    companies_html_body = "".join(companies_html_body)
    if len(analysis) == 0:
        analysis_html_body = f"""
        <body>
            <div>
                <p>
                    Bom dia,
                    <br/><br/>
                    The insurance payment analysis of {datetime.date.today().month} of {datetime.date.today().year}, of the companies in the attachments are correct.
                    <br/><br/>
                        <font color="red" >
                            No terminated employees were found invoiced.
                        </font>
                    <br/><br/>
                    Best regards,,
                    <br/><br/>
                    Caio Pavesi (Automação)
                <p/>
            <div/>
        </body>
        """
    else:
        analysis_html_body = f"""
        <body>
            <div>
                <p>
                    Bom dia,
                    <br/><br/>
                    Bellow the insurance payment analysis of {datetime.date.today().month} of {datetime.date.today().year}, regarding companies: {companies}.
                </p>
            </div>
            {companies_html_body}
            <div>
                <p>
                    Best regards,
                    <br/><br/>
                    Caio Pavesi (Automação)
                <p/>
            <div/>
        </body>
        """
    
    email = outlook_app.send_email(
        to = recipient
        ,subject = f"Insurance payment analysis, {datetime.date.today().month} of {datetime.date.today().year}"
        ,attachments = extracts
        ,html_body = analysis_html_body
    )
    Log().log(f"{datetime.datetime.now()} - 700000 [PRD] - Anaysis sent\n")
    
    return email


def clean_data(folders: list):
    """Removes all the content of a folder

    Args:
        folders (list): List of folders to clean

    Returns:
        deleted (list): All the deleted file names
    """
    
    deleted = list()
    
    for folder in folders:
        for file in os.listdir(folder):
            os.remove(folder / file)
            deleted.append(file)
    Log().log(f"{datetime.datetime.now()} - 700000 [PRD] - Temporary files removed\n")
    
    return deleted


# --- main --- #

def main():
    """Reads extract data from a Outlook (desktop app) folder and compares the employees retrieved from it with
    our Metabase employee data, the objective is to seek dismissed employees that we still pay insurance."""
    
    # Conig
    directory = config["Environment"]["Directory"]
    
    # Log setup
    Log.output_file =  directory.parent / "log.log"
    Log().log(f"{datetime.datetime.now()} - 700000 [PRD] - Process started\n")
    
    # Date vars
    today = datetime.datetime.today().date()
    monday = today - datetime.timedelta(days= today.weekday())
    sunday = monday + datetime.timedelta(days=6)
    
    # Outlook vars
    outlook_recipient_accounts = "company@email.com"
    outlook_search_account = "company@email.com"
    outlook_search_folder = "Nested/Folders/Names"
    outlook_output_folder = directory
    outlook_search_date = [monday, sunday]
    outlook_target_file = "_ATL.pdf"
    
    # Metabase vars
    metabase_username = credentials["Metabase"]["Username"]
    metabase_password = credentials["Metabase"]["Password"]
    metabase_domain = credentials["Metabase"]["Domain"]
    metabase_employee_table = 516
    metabase_rdb_database = 36
    
    # Obter o extrato de segurados p/empresa
    insurance_extracts = get_extracts(outlook_search_account, outlook_search_folder, outlook_output_folder, outlook_search_date, outlook_target_file)
    
    # ler os extratos
    extracts_content = read_extracts(insurance_extracts, outlook_output_folder)

    # Obter os colaboradores no metabase
    employees_data = get_employees(metabase_username, metabase_password, metabase_domain, metabase_rdb_database, metabase_employee_table)

    # Análise
    analysis_data = do_analysis(extracts_content, employees_data)

    # Enviar e-mail com os resultados
    email_subject = send_analysis(outlook_search_account, outlook_recipient_accounts, analysis_data, outlook_output_folder)

    # Eliminar arquivos desensessários
    deleted = clean_data([outlook_output_folder])

    Log().log(f"{datetime.datetime.now()} - 700000 [PRD] - Process completed successfully\n")
    
    return 0


# --- EXECUTABLE --- #

if __name__ == "__main__":
    main()
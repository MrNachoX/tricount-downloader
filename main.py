import requests
import json
import uuid
import os
from datetime import datetime
import rsa
import openpyxl
import csv
from tqdm import tqdm

class TricountAPI:
    def __init__(self):
        self.base_url = "https://api.tricount.bunq.com"
        self.app_installation_id = str(uuid.uuid4())
        self.public_key, self.private_key = rsa.newkeys(2048)
        self.rsa_public_key_pem = self.public_key.save_pkcs1(format="PEM").decode()
        self.headers = {
            "User-Agent": "com.bunq.tricount.android:RELEASE:7.0.7:3174:ANDROID:13:C",
            "app-id": self.app_installation_id,
            "X-Bunq-Client-Request-Id": "049bfcdf-6ae4-4cee-af7b-45da31ea85d0"
        }
        self.auth_token = None
        self.user_id = None

    def authenticate(self):
        auth_url = f"{self.base_url}/v1/session-registry-installation"
        auth_payload = {
            "app_installation_uuid": self.app_installation_id,
            "client_public_key": self.rsa_public_key_pem,
            "device_description": "Android"
        }
        response = requests.post(auth_url, json=auth_payload, headers=self.headers)
        response.raise_for_status()
        auth_data = response.json()

        response_items = auth_data["Response"]
        self.auth_token = next(item["Token"]["token"] for item in response_items if "Token" in item)
        self.user_id = next(item["UserPerson"]["id"] for item in response_items if "UserPerson" in item)
        self.headers["X-Bunq-Client-Authentication"] = self.auth_token

    def fetch_tricount_data(self, tricount_key):
        tricount_url = f"{self.base_url}/v1/user/{self.user_id}/registry?public_identifier_token={tricount_key}"
        response = requests.get(tricount_url, headers=self.headers)
        response.raise_for_status()
        return response.json()

class TricountHandler:
    @staticmethod
    def get_tricount_title(data):
        return data["Response"][0]["Registry"]["title"]

    @staticmethod
    def parse_tricount_data(data):
        registry = data["Response"][0]["Registry"]
        memberships = [
            {
                "Name": m["RegistryMembershipNonUser"]["alias"]["display_name"],
                "Balance": m["RegistryMembershipNonUser"]["balance"]["value"],
                "Currency": m["RegistryMembershipNonUser"]["balance"]["currency"]
            }
            for m in registry["memberships"]
        ]

        transactions = []
        for entry in registry["all_registry_entry"]:
            transaction = entry["RegistryEntry"]
            type_transaction = transaction["type_transaction"]
            who_paid = transaction["membership_owned"]["RegistryMembershipNonUser"]["alias"]["display_name"]
            total = float(transaction["amount"]["value"]) * -1
            currency = transaction["amount"]["currency"]
            description = transaction.get("description", "")
            when = transaction["date"]
            shares = {
                alloc["membership"]["RegistryMembershipNonUser"]["alias"]["display_name"]: abs(float(alloc["amount"]["value"]))
                for alloc in transaction["allocations"]
                }
            category = transaction["category"]
            attachments = transaction.get("attachment", [])

            transactions.append({
                "Type": type_transaction,
                "Who Paid": who_paid,
                "Total": total,
                "Currency": currency,
                "Description": description,
                "When": when,
                "Shares": shares,
                "Category": category,
                "Attachments": attachments
            })

        return memberships, transactions

    @staticmethod
    def download_attachments(transactions, download_folder):
        os.makedirs(download_folder, exist_ok=True)
        file_counter = 1
        total_files = sum(len(transaction["Attachments"]) for transaction in transactions)
        print(f"Total Attachments: {total_files}")

        if total_files == 0:
            return

        with tqdm(total=total_files, desc="Downloading attachments") as progress_bar:
            for transaction in transactions:
                attachment_files = []
                for attach in transaction["Attachments"]:
                    if "urls" in attach and attach["urls"]:
                        url = attach["urls"][0]["url"]
                        extension = os.path.splitext(url.split("?")[0])[1] or ".file"
                        file_name = f"receipt_{file_counter}{extension}"
                        file_path = os.path.join(download_folder, file_name)
                        TricountHandler.download_file(url, file_path)
                        attachment_files.append(file_name)
                        file_counter += 1
                        progress_bar.update(1)
                transaction["File Names"] = ", ".join(attachment_files)

    @staticmethod
    def download_file(url, file_path):
        response = requests.get(url)
        response.raise_for_status()
        with open(file_path, "wb") as file:
            file.write(response.content)

    @staticmethod
    def prepare_transaction_data(transaction):
        """
        Helper method to prepare the data for each transaction.
        Extracts involved people, formatted date, and attachment URLs.
        """
        # List of involved people involved in the transaction
        involved = ", ".join([name for name, amount in transaction["Shares"].items() if amount > 0])

        # Prepare the row data for the transaction
        row_data = [
            transaction["Who Paid"],
            transaction["Total"],
            transaction["Currency"],
            transaction["Description"],
            datetime.strptime(transaction["When"], "%Y-%m-%d %H:%M:%S.%f").strftime("%Y-%m-%d"),
            involved,
            transaction.get("File Names", ""),
            ", ".join([attach["urls"][0]["url"] for attach in transaction["Attachments"] if "urls" in attach and attach["urls"]])
        ]
        
        return row_data

    @staticmethod
    def prepare_sesterce_transaction_data(transaction, members):
        """
        Helper method to prepare the data for each transaction in the sesterce format.
        A row contains: 
        Date, Title, 
        Paid by Member A, Paid by Member B, ... , 
        Paid for Member A, Paid for Member B, ... ,
        Currency, Category
        """
        # create Paid by data
        paid_by = [0.0] * len(members)
        payer = transaction["Who Paid"]
        paid_by[members.index(payer)] = transaction["Total"]

        # create Paid for data
        paid_for = [0.0] * len(members)
        # paid_for_member is the name of the person that is involved in the transaction and didn't pay
        for paid_for_member, amount in transaction["Shares"].items():
            paid_for[members.index(paid_for_member)] = amount

        # Determine the category based on the transaction type
        type_transaction = transaction["Type"]
        category = ""  # Default empty

        if type_transaction == "BALANCE":
            category = "Money Transfer"
        elif type_transaction == "INCOME":
            # Negate paid_for values for income
            paid_for = [-amount for amount in paid_for]
            category = transaction["Category"] if transaction["Category"] != "UNCATEGORIZED" else ""
        elif type_transaction == "NORMAL":
            # Use the category if present
            category = transaction["Category"] if transaction["Category"] != "UNCATEGORIZED" else ""



        # Prepare the row data for the transaction
        row_data = [
            datetime.strptime(transaction["When"], "%Y-%m-%d %H:%M:%S.%f").strftime("%Y-%m-%d"),
            transaction["Description"],
            *paid_by,
            *paid_for,
            transaction["Currency"],
            category
        ]
        
        return row_data

    @staticmethod
    def write_to_excel(transactions, file_name):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Tricount Transactions"

        headers = ["Who Paid", "Total", "Currency", "Description", "When", "Involved", "File Names", "Attachment URLs"]
        sheet.append(headers)


        for transaction in transactions:
            row_data = TricountHandler.prepare_transaction_data(transaction)
            sheet.append(row_data)

        workbook.save(f"{file_name}.xlsx")
        print(f"Transactions have been saved to {file_name}.xlsx.")

    @staticmethod
    def write_to_csv(transactions, file_name):
        """
        Writes transaction data to a CSV file with the given file name.

        Parameters:
        - transactions (list): A list of transaction data.
        - file_name (str): The name of the CSV file to save the data to (without the .csv extension).

        The CSV file will have the following headers:
        "Who Paid", "Total", "Currency", "Description", "When", "Involved", "File Names", "Attachment URLs"

        Each transaction will be processed by the `prepare_transaction_data` method and written to the file.
        """
        with open(f"{file_name}.csv", "w") as csvfile:
            headers = ["Who Paid", "Total", "Currency", "Description", "When", "Involved", "File Names", "Attachment URLs"]
            transaction_writer = csv.writer(csvfile, delimiter=";")
            transaction_writer.writerow(headers)

            # Iterate through each transaction and write its data to the CSV file
            for transaction in transactions:
                row_data = TricountHandler.prepare_transaction_data(transaction)
                transaction_writer.writerow(row_data)

            print(f"Transactions have been saved to {file_name}.csv.")

    @staticmethod
    def write_to_sesterce_csv(memberships, transactions, file_name):
        """
        Writes transaction data in a specific format for Sesterce to a CSV file with the given file name.

        Parameters:
        - memberships (list): A list of memberships where each membership is a dictionary containing a "Name" key.
        - transactions (list): A list of transaction data.
        - file_name (str): The name of the CSV file to save the data to (without the .csv extension).

        The CSV file will have the following headers:
        "Date", "Title", "Paid by member" for each member, "Paid for member" for each member, "Currency", "Category"

        Each transaction will be processed by the `prepare_sesterce_transaction_data` method and written to the file.
        """
        # Sort members alphabetically based on their "Name"
        members = sorted([member["Name"] for member in memberships])

        with open(f"{file_name}.csv", "w") as csvfile:
            headers = ["Date", "Title"] + [f"Paid by {member}" for member in members] + [f"Paid for {member}" for member in members] + ["Currency", "Category"]
            transaction_writer = csv.writer(csvfile, delimiter=",")  # Sesterce expects "," delimiter
            transaction_writer.writerow(headers)

            # Iterate through each transaction and write its data to the CSV file
            for transaction in transactions:
                row_data = TricountHandler.prepare_sesterce_transaction_data(transaction, members)
                transaction_writer.writerow(row_data)

            print(f"Transactions have been saved to {file_name}.csv.")


if __name__ == "__main__":
    # example key
    tricount_key = "tISWyMCgrIMgFuxudZ"

    api = TricountAPI()
    api.authenticate()
    data = api.fetch_tricount_data(tricount_key)

    # save data to local file
    with open('response_data.json', 'w') as f:
        json.dump(data, f, indent=2)

    # load data from local file
    #with open('response_data.json', 'r') as f:
    #    data = json.load(f)

    handler = TricountHandler()
    tricount_title = handler.get_tricount_title(data)

    memberships, transactions = handler.parse_tricount_data(data)

    handler.write_to_csv(transactions, file_name=f"Transactions {tricount_title}")

    #handler.write_to_excel(transactions, file_name=f"Transactions {tricount_title}")
    #handler.write_to_sesterce_csv(memberships, transactions, f"Transaction {tricount_title} (Sesterce)")
    #handler.download_attachments(transactions, download_folder=f"Attachments {tricount_title}")

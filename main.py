import pandas as pd
from openpyxl import load_workbook
from pathlib import Path
import datetime

# Define the Library class with Excel functionalities
class Library:
    def __init__(self, listofBooks, excel_filename="borrowed_books.xlsx"):
        self.books = listofBooks
        self.excel_filename = excel_filename
        
        # Create an Excel file with the necessary columns if it doesn't exist
        if not Path(self.excel_filename).is_file():
            df = pd.DataFrame(columns=["Borrower Name", "Book Name", "Date Borrowed"])
            df.to_excel(self.excel_filename, index=False)

    def displayAvailableBooks(self):
        print(f"\n{len(self.books)} AVAILABLE BOOKS:")
        for book in self.books:
            print(" ♦-- " + book)
        print("\n")

    def borrowBook(self, name, bookname):
        if bookname not in self.books:
            print(
                f"{bookname} BOOK IS NOT AVAILABLE, EITHER TAKEN BY SOMEONE ELSE. WAIT UNTIL IT'S RETURNED.\n"
            )
        else:
            # Append to the track list
            track.append({name: bookname})
            print("BOOK ISSUED: THANK YOU, KEEP IT WITH CARE, AND RETURN ON TIME.\n")
            self.books.remove(bookname)

            # Add to the Excel sheet
            df = pd.read_excel(self.excel_filename)
            new_entry = {"Borrower Name": name, "Book Name": bookname, "Date Borrowed": pd.Timestamp.now()}
            df = df.append(new_entry, ignore_index=True)
            df.to_excel(self.excel_filename, index=False)

    def returnBook(self, bookname):
        print("BOOK RETURNED: THANK YOU!\n")
        self.books.append(bookname)

        # Remove from the Excel sheet
        df = pd.read_excel(self.excel_filename)
        book_returned = df[df["Book Name"] == bookname].index
        if not book_returned.empty:
            df = df.drop(book_returned)
            df.to_excel(self.excel_filename, index=False)

    def donateBook(self, bookname):
        print("BOOK DONATED: THANK YOU VERY MUCH, HAVE A GREAT DAY AHEAD.\n")
        self.books.append(bookname)


# Define the Student class
class Student:
    def requestBook(self):
        print("You want to borrow a book!")
        self.book = input("Enter the name of the book you want to borrow: ")
        return self.book

    def returnBook(self):
        print("You want to return a book!")
        name = input("Enter your name: ")
        self.book = input("Enter the name of the book you want to return: ")
        if {name: self.book} in track:
            track.remove({name: self.book})
        return self.book

    def donateBook(self):
        print("You want to donate a book!")
        self.book = input("Enter the name of the book you want to donate: ")
        return self.book


# Main code with Excel functionalities
if __name__ == "__main__":
    Chennailibrary= Library(
        ["vistas", "invention", "rich&poor", "indian", "macroeconomics", "microeconomics"]
    )
    student = Student()
    track = []

    print("\t\t\t\t\t♦♦♦♦♦♦♦ WELCOME TO THE CHENNAI LIBRARY ♦♦♦♦♦♦♦\n")
    print(
        """CHOOSE WHAT YOU WANT TO DO:-\n1. Listing all books\n2. Borrow books\n3. Return books\n4. Donate books\n5. Track books\n6. Exit the library\n"""
    )

    while True:
        try:
            usr_response = int(input("Enter your choice: "))
            if usr_response == 1:  # listing
                Delhilibrary.displayAvailableBooks()
            elif usr_response == 2:  # borrow
                Delhilibrary.borrowBook(
                    input("Enter your name: "), student.requestBook()
                )
            elif usr_response == 3:  # return
                Delhilibrary.returnBook(student.returnBook())
            elif usr_response == 4:  # donate
                Delhilibrary.donateBook(student.donateBook())
            elif usr_response == 5:  # track
                for i in track:
                    for key, value in i.items():
                        print(f"{value} book is taken/issued by {key}.")
                if len(track) == 0:
                    print("NO BOOKS ARE ISSUED!\n")
            elif usr_response == 6:  # exit
                print("THANK YOU!\n")
                break
            else:
                print("INVALID INPUT!\n")
        except Exception as e:
            print(f"{e}---> INVALID INPUT!\n")

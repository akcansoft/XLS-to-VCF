# akcanSoft XLS-to-VCF

This Excel file exports contact information from a table to a standard `.vcf` (vCard) file format. It allows you to easily import your contacts in bulk to your phonebook or email client.

![](https://1.bp.blogspot.com/-4aoyHryb96Y/Xc_FwMklOhI/AAAAAAACMSY/sHv1hpAWxrMLI05WAa7ZbpHxkkQZPa11ACLcBGAsYHQ/s1600/xlstovcf.png)

## 📥 Installation

1.  **Download the File:** Download the `akcanSoft XLS to VCF v1.0.xlsm` file from this repository to your computer.
2.  **Enable Macros:** When you open the file in Excel, a yellow security warning bar may appear at the top. Click the **"Enable Content"** or **"Enable Macros"** button to allow the VBA code to run.

## 📖 How to Use

1.  **Enter Your Data:** The Excel sheet is pre-configured with the following columns. Enter your contact information into the relevant rows, starting from row 2 (row 1 contains the headers).

    | Column | Field     | Description                    |
    |--------|-----------|--------------------------------|
    | A      | Name      | Contact's first name           |
    | B      | Surname   | Contact's last name            |
    | C      | Tel       | Contact's phone number         |
    | D      | Email     | Contact's email address        |

2.  **Run the Macro:**
    *   Go to the **"Developer"** tab in the Excel ribbon. (If you don't see this tab, you can enable it from File -> Options -> Customize Ribbon).
    *   Click the **"Macros"** button.
    *   Select the macro named **`Create_vCard`** from the list and click **"Run"**.
3.  **Save the VCF File:** A dialog box will appear, prompting you to choose where to save the generated `.vcf` file and what name to give it. Select your desired location and click "Save".

That's it! Your `.vcf` file will be created in the chosen folder, containing all the contacts from your Excel sheet. You can now transfer this file to your phone or other devices and import the contacts.

## ⚙️ Technical Details

The project is written entirely in **VBA (Visual Basic for Applications)** . The core logic is contained within the `Create_vCard.bas` module.

### How the Code Works (`Create_vCard.bas`)

The VBA macro automates the entire process:

1.  **Initialization:** It first checks for an existing `FileDialog` object to let the user choose the save location. It also defines the required vCard headers (`BEGIN:VCARD`, `VERSION:3.0`) and footer (`END:VCARD`).
2.  **Data Looping:** The code loops through each row of data in the "Sheet1" worksheet, starting from row 2 until it finds an empty cell in column A (Name).
3.  **vCard Formatting:** For each row with data, it builds a vCard entry. The macro structures the information according to the vCard 3.0 specification:
    *   `N:` and `FN:` properties are constructed using the Name and Surname columns.
    *   `TEL;TYPE=CELL:` property is created from the Tel column.
    *   `EMAIL;TYPE=INTERNET;TYPE=WORK;TYPE=pref:` property is created from the Email column.
4.  **File Writing:** All the generated vCard entries are written sequentially into the single `.vcf` file selected by the user. Each contact entry is separated by a blank line as per the standard.

## Contacts

Mesut Akcan  
**Blog:** https://akcanSoft.blogspot.com | https://mesutakcan.blogspot.com  
**YouTube:** www.youtube.com/mesutakcan

## ⭐ Support & Feedback

If you find this tool useful, please consider giving the project a **star** on GitHub! For feedback or issues, feel free to open an issue in the repository.

from tkcalendar import Calendar
import tkinter.font as tkFont
from tkinter import ttk
from tkinter import *
import ctypes
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch
from reportlab.lib import colors
from datetime import datetime


class PDF:
    def __init__(self):
        self.styles = getSampleStyleSheet()
        self.elements = []
        self.width, self.height = A4

    def add_company_info(self, name, address, phone, website):
        company_name_style = ParagraphStyle(
            "CompanyName",
            parent=self.styles["Title"],
            fontSize=24,
            spaceAfter=12,
            alignment=0,  # Left alignment
        )
        self.elements.append(Paragraph(name, company_name_style))
        self.elements.append(
            Paragraph(address, ParagraphStyle(name="CAddress", alignment=0))
        )
        self.elements.append(
            Paragraph(phone, ParagraphStyle(name="CPhone", alignment=0))
        )
        self.elements.append(
            Paragraph(
                f"<a href='{website}' color='black'>{website}</a>",
                ParagraphStyle(name="CLink", alignment=0),
            )
        )
        self.elements.append(Spacer(1, 0.5 * inch))

    def add_section_heading(self, heading):
        heading_style = ParagraphStyle(
            "Heading1",
            parent=self.styles["Heading1"],
            alignment=0,
            textColor=colors.white,
            backColor=colors.dimgrey,
            borderPadding=0.10 * inch,
        )
        self.elements.append(Paragraph(heading, heading_style))
        # self.elements.append(Spacer(1, 0.1 * inch))

    def create_section_heading(self, heading):
        heading_style = ParagraphStyle(
            "Heading2",
            parent=self.styles["Heading1"],
            alignment=0,
            fontSize=14,
            textColor=colors.white,
            backColor=colors.dimgrey,
            borderPadding=0.05 * inch,
        )

        return Paragraph(heading, heading_style)

    def add_customer_info(self, customer_info, sale_order_info):
        # Create nested tables for section headings
        parent_table_head_data = [
            [
                self.create_section_heading("Customer Information"),
                self.create_section_heading("Sales Order Information"),
            ]
        ]
        parent_table_head = Table(
            parent_table_head_data, colWidths=[self.width * 0.6, self.width * 0.4]
        )

        # Apply style to the parent table headings
        parent_table_head.setStyle(
            TableStyle(
                [
                    ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("GRID", (0, 0), (-1, -1), 1, colors.transparent),
                    ("LEFTPADDING", (0, 0), (0, 0), inch / 2),
                    ("RIGHTPADDING", (-1, -1), (-1, -1), inch / 2),
                ]
            )
        )

        # Create nested tables for customer info and sales order info
        parent_table_data = [
            [
                self.create_customer_info_table(customer_info),
                self.create_sales_order_info_table(sale_order_info),
            ]
        ]
        parent_table = Table(
            parent_table_data, colWidths=[self.width * 0.6, self.width * 0.4]
        )

        # Apply style to the parent table
        parent_table.setStyle(
            TableStyle(
                [
                    ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("GRID", (0, 0), (-1, -1), 1, colors.transparent),
                    ("LEFTPADDING", (0, 0), (0, 0), inch / 2.1),
                    ("RIGHTPADDING", (-1, -1), (-1, -1), inch / 2.1),
                ]
            )
        )

        # Add elements to the document
        self.elements.append(parent_table_head)
        self.elements.append(parent_table)
        self.elements.append(Spacer(1, 0.07 * inch))

    def create_customer_info_table(self, customer_info):
        # Create table for customer information
        customer_data = [
            [Paragraph(k, self.styles["Heading3"]), Paragraph(v, self.styles["Normal"])]
            for k, v in customer_info.items()
        ]
        customer_table = Table(
            customer_data,
            colWidths=[None, None],
            style=[
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("GRID", (0, 0), (-1, -1), 1, colors.black),
            ],
        )
        return customer_table

    def create_sales_order_info_table(self, sale_order_info):
        # Create table for sales order information
        sale_order_data = [
            [Paragraph(k, self.styles["Heading3"]), Paragraph(v, self.styles["Normal"])]
            for k, v in sale_order_info.items()
        ]
        sale_order_table = Table(
            sale_order_data,
            colWidths=[None, None],
            style=[
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("GRID", (0, 0), (-1, -1), 1, colors.black),
            ],
        )
        return sale_order_table

    def add_sales_person(self, sales_person):
        self.add_section_heading("Sales Person")

        styles = getSampleStyleSheet()
        normal_style = styles["Normal"]

        sales_person_data = [
            ["Employee ID", "Employee Name", "Designation", "Department"]
        ] + [
            [Paragraph(str(cell), normal_style) for cell in row] for row in sales_person
        ]

        sales_person_table = Table(
            sales_person_data,
            colWidths=[
                (self.width / 4) - (inch / 4),
                (self.width / 4) - (inch / 4),
                (self.width / 4) - (inch / 4),
                (self.width / 4) - (inch / 4),
            ],
            style=[
                ("GRID", (0, 0), (-1, -1), 1, colors.black),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ],
        )
        self.elements.append(sales_person_table)
        self.elements.append(Spacer(1, 0.12 * inch))

    def add_payment_schedule(self, payment_schedule):
        self.add_section_heading("Payment Schedule/Term")

        styles = getSampleStyleSheet()
        normal_style = styles["Normal"]

        payment_schedule_data = [
            [
                "Description",
                "Estimated Date",
                "Actual Date",
                "Per%",
                "Amount",
                "Pay Method",
            ]
        ] + [
            [Paragraph(str(cell), normal_style) for cell in row]
            for row in payment_schedule
        ]

        # Adjusted column widths to fit within the page
        col_widths = [
            (self.width * 0.25) - (inch / 6),  # Description
            (self.width * 0.15) - (inch / 6),  # Estimated Date
            (self.width * 0.15) - (inch / 6),  # Actual Date
            (self.width * 0.10) - (inch / 6),  # Per%
            (self.width * 0.15) - (inch / 6),  # Amount
            (self.width * 0.20) - (inch / 6),  # Pay Method
        ]

        payment_schedule_table = Table(
            payment_schedule_data,
            colWidths=col_widths,
            style=[
                ("GRID", (0, 0), (-1, -1), 1, colors.black),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ],
        )
        self.elements.append(payment_schedule_table)
        self.elements.append(Spacer(1, 0.12 * inch))

    def add_details(self, details):
        self.add_section_heading("Details")

        # Define column widths based on your layout requirements
        col_widths = [
            self.width * 0.06,  # Sno
            self.width * 0.22,  # Item Description
            self.width * 0.10,  # Quantity
            self.width * 0.10,  # Unit Price
            self.width * 0.10,  # Gross Amount
            self.width * 0.10,  # Discount Per%
            self.width * 0.10,  # Discount Amount
            self.width * 0.10,  # Net Amount
        ]

        # Create a sample style for table content
        styles = getSampleStyleSheet()
        normal_style = styles["Normal"]

        # Wrap text using Paragraph for each cell
        details_data = [
            [Paragraph(str(cell), normal_style) for cell in row] for row in details
        ]

        details_table = Table(
            details_data,
            colWidths=col_widths,
            style=[
                ("GRID", (0, 0), (-1, -1), 1, colors.black),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ],
            # minRowHeights can be adjusted based on your content needs
            minRowHeights=[inch * 0.4] * len(details_data),
        )

        self.elements.append(details_table)
        self.elements.append(Spacer(1, 0.50 * inch))

    def add_signatures(self):
        # Sample signatures (can be replaced with any content)
        data = [
            ["______________", "______________", "______________"],
            [
                "     Prepared by",
                "Authorized by",
                "Received by    ",
            ],
        ]

        # Create a table with one row and multiple columns
        table = Table(
            data,
            colWidths=[
                (self.width / 3) - (inch / 3),
                (self.width / 3) - (inch / 3),
                (self.width / 3) - (inch / 3),
            ],
        )  # Adjust column widths as needed

        # Apply style to the table for transparent borders
        table.setStyle(
            TableStyle(
                [
                    ("ALIGN", (0, 0), (0, 1), "LEFT"),
                    ("ALIGN", (1, 0), (1, 1), "CENTER"),
                    ("ALIGN", (2, 0), (2, 1), "RIGHT"),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("TEXTCOLOR", (0, 0), (-1, -1), colors.black),
                    (
                        "LINEBELOW",
                        (0, 0),
                        (-1, -1),
                        0,
                        colors.transparent,
                    ),  # Transparent border
                ]
            )
        )

        self.elements.append(table)

    def add_footer(self, canvas, doc):
        canvas.saveState()
        footer_text = "Developed By Al-Feroz Tech"
        datetime_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        canvas.drawString(inch / 2, 0.3 * inch, f"Page {doc.page}")
        canvas.drawString((doc.width / 2) - (inch / 1.5), 0.3 * inch, footer_text)
        canvas.drawString(doc.width - inch, 0.3 * inch, datetime_text)
        canvas.restoreState()

    def generate(
        self,
        company_info,
        customer_info,
        sale_order_info,
        sales_person,
        payment_schedule,
        details,
    ):
        # Format the datetime to avoid invalid characters
        current_datetime = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename = f"{company_info['name']} - {current_datetime}.pdf"

        doc = SimpleDocTemplate(
            filename,
            pagesize=A4,
            rightMargin=(inch / 2),
            leftMargin=(inch / 2),
            topMargin=(inch / 2),
            bottomMargin=(inch / 2),
        )

        doc.author = "Al-Feroz Tech"
        doc.title = company_info["name"]
        doc.subject = f"Invoice: {company_info['name']} - {current_datetime}"

        self.add_company_info(**company_info)
        self.add_customer_info(customer_info, sale_order_info)
        self.add_sales_person(sales_person)
        self.add_payment_schedule(payment_schedule)
        self.add_details(details)
        self.add_signatures()

        doc.build(
            self.elements, onFirstPage=self.add_footer, onLaterPages=self.add_footer
        )


# Creating class AutoScrollbar
class AutoScrollbar(Scrollbar):
    # A scrollbar that hides itself if it's not needed. Only works if you use the grid geometry manager.
    def set(self, lo, hi):
        if float(lo) <= 0.0 and float(hi) >= 1.0:
            self.grid_remove()
        else:
            self.grid()
        Scrollbar.set(self, lo, hi)

    def pack(self, **kw):
        raise TclError("cannot use pack with this widget")

    def place(self, **kw):
        raise TclError("cannot use place with this widget")


class InvoiceGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("AL-FEROZ TECH - Invoice Generator")
        self.root.minsize(1200, 500)
        self.root.resizable(True, True)
        self.root.state("zoomed")

        self.icon_path = "favicon_io/favicon.ico"
        try:
            self.root.iconbitmap(True, self.icon_path)
        except:
            self.icon_path = "favicon.ico"
            self.root.iconbitmap(True, self.icon_path)

        myappid = "ALFEROZTECH.invoicegenerator.version"
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

        self.currentDate = Calendar(self.root).get_date()

        self.rootMain()

    def pick_date(self, datePickedVar, datePickedBTN):
        dateWin = Tk()
        # dateWin.geometry("500x300")
        dateWin.iconbitmap(TRUE, self.icon_path)
        dateWin.title("Pick A Date")

        cal = Calendar(dateWin)
        cal.pack(fill=X)

        pick_btn = Button(
            dateWin,
            text="Done",
            background="black",
            fg="white",
            command=lambda: self.remove_date(
                dateWin, cal, datePickedVar, datePickedBTN
            ),
        )
        pick_btn.pack(fill=X)

        dateWin.focus()
        dateWin.mainloop()

    def remove_date(self, dateWin, cal, datePickedVar, datePickedBTN):
        selected_date = cal.get_date()
        datePickedVar.set(selected_date)
        datePickedBTN.configure(text="Edit")
        dateWin.destroy()

    def rootMain(self):
        # Defining vertical scrollbar
        verscrollbar = AutoScrollbar(self.root, orient=VERTICAL)
        verscrollbar.grid(row=0, column=1, sticky=N + S)

        # Defining horizontal scrollbar
        horiscrollbar = AutoScrollbar(self.root, orient=HORIZONTAL)
        horiscrollbar.grid(row=1, column=0, sticky=E + W)

        # Creating scrolled canvas
        canvas = Canvas(
            self.root, yscrollcommand=verscrollbar.set, xscrollcommand=horiscrollbar.set
        )
        canvas.grid(row=0, column=0, sticky=N + S + E + W)

        verscrollbar.config(command=canvas.yview)
        horiscrollbar.config(command=canvas.xview)

        # Making the canvas expandable
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        # creating canvas contents
        main_frame = Frame(canvas)
        main_frame.grid(sticky=N + S + E + W)

        self.mainHead(main_frame)
        self.mainBody(main_frame)
        self.mainFooter(main_frame)

        # Creating canvas window
        canvas.create_window((0, 0), anchor=NW, window=main_frame)

        # Configuring canvas
        main_frame.update_idletasks()
        canvas.config(scrollregion=canvas.bbox("all"))

    def mainHead(self, mainFrame):
        headFrame = Frame(mainFrame, name="headframe")
        headFrame.pack(fill=X, side=TOP)

        Mainheading = Label(
            headFrame,
            text="Invoice Generator",
            fg="black",
            font=("Sarif", 20, tkFont.BOLD),
            pady=30,
        )
        Mainheading.pack(fill=X, side=TOP)

    def mainBody(self, mainFrame):
        bodyFrame = Frame(mainFrame, name="bodyframe")
        bodyFrame.pack(fill=X, padx=50, ipadx=80)

        self.Section1 = Frame(bodyFrame, name="section1")
        self.Section1.pack(fill=X)

        self.Section2 = Frame(bodyFrame, name="section2")
        self.Section2.pack(fill=X)

        self.bodySubSection1 = Frame(self.Section1, name="bodysubsection1")
        self.bodySubSection1.pack(fill=X, side=LEFT)

        self.bodySubSection2 = Frame(self.Section1, name="bodysubsection2")
        self.bodySubSection2.pack(fill=X, side=RIGHT)

        self.companySection()
        self.salesOrderInfoSection()
        self.paymentScheduleSection()
        self.salesPersonSection()
        self.customerSection()
        self.detailSection()

    def companySection(self):
        CompanySection = LabelFrame(
            self.bodySubSection1,
            name="companysection",
            text="Company Information",
            fg="black",
            font=("Sarif", 16, tkFont.BOLD),
        )
        CompanySection.pack(fill=X)

        CompanyEntrySection = Frame(CompanySection, name="companyentrysection")
        CompanyEntrySection.pack(fill=X, pady=20)

        CompanyNameLabel = Label(
            CompanyEntrySection, text="Company Name:", font=("Sarif", 10, tkFont.BOLD)
        )
        CompanyNameLabel.pack(fill=X, side=LEFT, ipadx=5)

        self.CompanyNameinputField = Entry(
            CompanyEntrySection,
            textvariable="CompanyName",
            cursor="hand2",
            justify=CENTER,
            relief=SOLID,
        )
        self.CompanyNameinputField.pack(fill=X, side=LEFT, padx=10)

        CompanyLinkLabel = Label(
            CompanyEntrySection, text="Company Link:", font=("Sarif", 10, tkFont.BOLD)
        )
        CompanyLinkLabel.pack(fill=X, side=LEFT, padx=5)

        self.CompanyLinkinputField = Entry(
            CompanyEntrySection,
            textvariable="CL",
            cursor="hand2",
            justify=CENTER,
            relief=SOLID,
        )
        self.CompanyLinkinputField.pack(fill=X, side=LEFT)

        CompanyAddressLabel = Label(
            CompanySection, text="Address:", font=("Sarif", 10, tkFont.BOLD)
        )
        CompanyAddressLabel.pack(fill=X, side=LEFT, ipadx=5)

        self.CompanyAddressinputField = Text(
            CompanySection,
            width=50,
            height=3,
            cursor="hand2",
            relief=SOLID,
            font=("Sarif", 12, tkFont.NORMAL),
        )
        self.CompanyAddressinputField.pack(fill=X, side=LEFT, padx=10, pady=10)

    def salesOrderInfoSection(self):
        OrderInformationSection = LabelFrame(
            self.bodySubSection1,
            name="orderinformationsection",
            text="Sales Order Information",
            fg="black",
            font=("Sarif", 16, tkFont.BOLD),
        )
        OrderInformationSection.pack(fill=X, pady=30)

        So_ID_Date_Frame = Frame(OrderInformationSection, name="soidframe")
        So_ID_Date_Frame.pack(fill=X, pady=10)

        SoIDLabel = Label(
            So_ID_Date_Frame, text="SO #:", font=("Sarif", 10, tkFont.BOLD)
        )
        SoIDLabel.pack(fill=X, side=LEFT, ipadx=5)

        self.SoIDinputField = Entry(
            So_ID_Date_Frame,
            textvariable="SOID",
            cursor="hand2",
            justify=CENTER,
            relief=SOLID,
        )
        self.SoIDinputField.pack(fill=X, side=LEFT, padx=10)

        SODateLabel = Label(
            So_ID_Date_Frame, text="SO Date:", font=("Sarif", 10, tkFont.BOLD)
        )
        SODateLabel.pack(fill=X, side=LEFT, ipadx=10)

        self.SODatePickedVar = StringVar(value=self.currentDate)
        SODatePicked = Label(
            So_ID_Date_Frame,
            textvariable=self.SODatePickedVar,
            font=("Sarif", 10, tkFont.BOLD),
        )
        SODatePicked.pack(fill=X, side=LEFT)

        SODate = Button(
            So_ID_Date_Frame,
            text="Pick Date",
            command=lambda: self.pick_date(self.SODatePickedVar, SODate),
        )
        SODate.pack(fill=X, side=LEFT)

        DDate_Qref_Frame = Frame(OrderInformationSection, name="deliverydatesection")
        DDate_Qref_Frame.pack(fill=X, pady=5)

        QRefLabel = Label(
            DDate_Qref_Frame, text="Qoutation Ref:", font=("Sarif", 10, tkFont.BOLD)
        )
        QRefLabel.pack(fill=X, side=LEFT, ipadx=5)

        self.QRefinputField = Entry(
            DDate_Qref_Frame,
            textvariable="QRef",
            cursor="hand2",
            justify=CENTER,
            relief=SOLID,
        )
        self.QRefinputField.pack(fill=X, side=LEFT, padx=10)

        DeliveryDateLabel = Label(
            DDate_Qref_Frame, text="Delivery Date:", font=("Sarif", 10, tkFont.BOLD)
        )
        DeliveryDateLabel.pack(fill=X, side=LEFT, ipadx=10)

        self.DeliveryDatePickedVar = StringVar(value=self.currentDate)
        DeliveryDatePicked = Label(
            DDate_Qref_Frame,
            textvariable=self.DeliveryDatePickedVar,
            font=("Sarif", 10, tkFont.BOLD),
        )
        DeliveryDatePicked.pack(fill=X, side=LEFT)

        DeliveryDate = Button(
            DDate_Qref_Frame,
            text="Pick Date",
            command=lambda: self.pick_date(self.DeliveryDatePickedVar, DeliveryDate),
        )
        DeliveryDate.pack(fill=X, side=LEFT)

    def paymentScheduleSection(self):
        TermSection = LabelFrame(
            self.bodySubSection1,
            name="termsection",
            text="Payment Schedule / Term",
            fg="black",
            font=("Sarif", 16, tkFont.BOLD),
        )
        TermSection.pack(fill=X)

        Desc_PM_Frame = Frame(TermSection, name="descandpmframe")
        Desc_PM_Frame.pack(fill=X, pady=10)

        DescLabel = Label(
            Desc_PM_Frame, text="Description:", font=("Sarif", 10, tkFont.BOLD)
        )
        DescLabel.pack(fill=X, side=LEFT, ipadx=5)

        self.DescinputField = Entry(
            Desc_PM_Frame,
            textvariable="Desc",
            cursor="hand2",
            justify=CENTER,
            relief=SOLID,
        )
        self.DescinputField.pack(fill=X, side=LEFT, padx=10)

        payMethodLabel = Label(
            Desc_PM_Frame, text="Pay Method:", font=("Sarif", 10, tkFont.BOLD)
        )
        payMethodLabel.pack(fill=X, side=LEFT, ipadx=5)

        self.payMethodinputField = Entry(
            Desc_PM_Frame,
            textvariable="PM",
            cursor="hand2",
            justify=CENTER,
            relief=SOLID,
        )
        self.payMethodinputField.pack(fill=X, side=LEFT, padx=10)

        Per_Amount_Frame = Frame(TermSection, name="perandamountframe")
        Per_Amount_Frame.pack(fill=X, pady=10)

        PerLabel = Label(
            Per_Amount_Frame, text="Per%:", font=("Sarif", 10, tkFont.BOLD)
        )
        PerLabel.pack(fill=X, side=LEFT, ipadx=5)

        self.PerinputField = Entry(
            Per_Amount_Frame,
            textvariable="Per",
            cursor="hand2",
            justify=CENTER,
            relief=SOLID,
        )
        self.PerinputField.pack(fill=X, side=LEFT, padx=10)

        AmountLabel = Label(
            Per_Amount_Frame, text="Amount:", font=("Sarif", 10, tkFont.BOLD)
        )
        AmountLabel.pack(fill=X, side=LEFT, ipadx=5)

        self.AmountinputField = Entry(
            Per_Amount_Frame,
            textvariable="Amount",
            cursor="hand2",
            justify=CENTER,
            relief=SOLID,
        )
        self.AmountinputField.pack(fill=X, side=LEFT, padx=10)

        Est_AD_Frame = Frame(TermSection, name="estandactframe")
        Est_AD_Frame.pack(fill=X, pady=10)

        EstDateLabel = Label(
            Est_AD_Frame, text="Estimated Date:", font=("Sarif", 10, tkFont.BOLD)
        )
        EstDateLabel.pack(fill=X, side=LEFT, ipadx=5)

        self.EstDatePickedVar = StringVar(value=self.currentDate)
        EstDatePicked = Label(
            Est_AD_Frame,
            textvariable=self.EstDatePickedVar,
            font=("Sarif", 10, tkFont.BOLD),
        )
        EstDatePicked.pack(fill=X, side=LEFT, ipadx=5)

        EstDate = Button(
            Est_AD_Frame,
            text="Pick Date",
            command=lambda: self.pick_date(self.EstDatePickedVar, EstDate),
        )
        EstDate.pack(fill=X, side=LEFT)

        ActDateLabel = Label(
            Est_AD_Frame, text="Actual Date:", font=("Sarif", 10, tkFont.BOLD)
        )
        ActDateLabel.pack(fill=X, side=LEFT, ipadx=5)

        self.ActDatePickedVar = StringVar(value=self.currentDate)
        ActDatePicked = Label(
            Est_AD_Frame,
            textvariable=self.ActDatePickedVar,
            font=("Sarif", 10, tkFont.BOLD),
        )
        ActDatePicked.pack(fill=X, side=LEFT, ipadx=5)

        ActDate = Button(
            Est_AD_Frame,
            text="Pick Date",
            command=lambda: self.pick_date(self.ActDatePickedVar, ActDate),
        )
        ActDate.pack(fill=X, side=LEFT)

    def salesPersonSection(self):
        SalesPersonSection = LabelFrame(
            self.bodySubSection2,
            name="salespersonsection",
            text="Sales Person",
            fg="black",
            font=("Sarif", 16, tkFont.BOLD),
            padx=10,
            pady=20,
        )
        SalesPersonSection.pack(fill=X, pady=25)

        Emp_ID_N_Frame = Frame(SalesPersonSection, name="empidandnameframe")
        Emp_ID_N_Frame.pack(fill=X, pady=10)

        EmpIDLabel = Label(
            Emp_ID_N_Frame, text="Employee ID:", font=("Sarif", 10, tkFont.BOLD)
        )
        EmpIDLabel.pack(fill=X, side=LEFT, ipadx=5)

        self.EmpIDinputField = Entry(
            Emp_ID_N_Frame,
            textvariable="EmpID",
            cursor="hand2",
            justify=CENTER,
            relief=SOLID,
        )
        self.EmpIDinputField.pack(fill=X, side=LEFT, padx=10)

        EmpNameLabel = Label(
            Emp_ID_N_Frame, text="Employee Name:", font=("Sarif", 10, tkFont.BOLD)
        )
        EmpNameLabel.pack(fill=X, side=LEFT, ipadx=5)

        self.EmpNameinputField = Entry(
            Emp_ID_N_Frame,
            textvariable="EName",
            cursor="hand2",
            justify=CENTER,
            relief=SOLID,
        )
        self.EmpNameinputField.pack(fill=X, side=LEFT, padx=10)

        EmpDesignationFrame = Frame(SalesPersonSection, name="empdesignationframe")
        EmpDesignationFrame.pack(fill=X, pady=10)

        EmpDesignationLabel = Label(
            EmpDesignationFrame,
            text="Employee Designation:",
            font=("Sarif", 10, tkFont.BOLD),
        )
        EmpDesignationLabel.pack(fill=X, side=LEFT, ipadx=5)

        self.EmpDesignationinputField = Entry(
            EmpDesignationFrame,
            textvariable="EmpDesignation",
            cursor="hand2",
            justify=CENTER,
            relief=SOLID,
        )
        self.EmpDesignationinputField.pack(fill=X, side=LEFT, padx=10)

        EmpDepartmentFrame = Frame(SalesPersonSection, name="empdepartmentframe")
        EmpDepartmentFrame.pack(fill=X, pady=10)

        EmpDepartmentLabel = Label(
            EmpDepartmentFrame, text="Department:", font=("Sarif", 10, tkFont.BOLD)
        )
        EmpDepartmentLabel.pack(fill=X, side=LEFT, ipadx=5)

        self.EmpDepartmentinputField = Entry(
            EmpDepartmentFrame,
            textvariable="EmpDep",
            cursor="hand2",
            justify=CENTER,
            relief=SOLID,
        )
        self.EmpDepartmentinputField.pack(fill=X, side=LEFT, padx=10)

    def customerSection(self):
        CustomerSection = LabelFrame(
            self.bodySubSection2,
            name="customersection",
            text="Customer Information",
            fg="black",
            font=("Sarif", 16, tkFont.BOLD),
            padx=10,
            pady=20,
        )
        CustomerSection.pack(fill=X, pady=25)

        Customer_N_E_Section = Frame(
            CustomerSection, name="customernameandemailsection"
        )
        Customer_N_E_Section.pack(fill=X)

        CustomerNameLabel = Label(
            Customer_N_E_Section, text="Customer Name:", font=("Sarif", 10, tkFont.BOLD)
        )
        CustomerNameLabel.pack(fill=X, side=LEFT, ipadx=5)

        self.CustomerNameinputField = Entry(
            Customer_N_E_Section,
            textvariable="CustomerName",
            cursor="hand2",
            justify=CENTER,
            relief=SOLID,
        )
        self.CustomerNameinputField.pack(fill=X, side=LEFT, padx=10)

        CustomerEmailLabel = Label(
            Customer_N_E_Section,
            text="Customer Email:",
            font=("Sarif", 10, tkFont.BOLD),
        )
        CustomerEmailLabel.pack(fill=X, side=LEFT, padx=5)

        self.CustomerEmailinputField = Entry(
            Customer_N_E_Section,
            textvariable="CEmail",
            cursor="hand2",
            justify=CENTER,
            relief=SOLID,
        )
        self.CustomerEmailinputField.pack(fill=X, side=LEFT)

        Customer_P_M_Section = Frame(
            CustomerSection, name="customerphoneandmobilesection"
        )
        Customer_P_M_Section.pack(fill=X, pady=5)

        CustomerPhoneLabel = Label(
            Customer_P_M_Section, text="Phone Number:", font=("Sarif", 10, tkFont.BOLD)
        )
        CustomerPhoneLabel.pack(fill=X, side=LEFT, ipadx=5)

        self.CustomerPhoneinputField = Entry(
            Customer_P_M_Section,
            textvariable="CustomerPN",
            cursor="hand2",
            justify=CENTER,
            relief=SOLID,
        )
        self.CustomerPhoneinputField.pack(fill=X, side=LEFT, padx=10)

        CustomerMobileLabel = Label(
            Customer_P_M_Section, text="Mobile Number:", font=("Sarif", 10, tkFont.BOLD)
        )
        CustomerMobileLabel.pack(fill=X, side=LEFT, padx=5)

        self.CustomerMobileinputField = Entry(
            Customer_P_M_Section,
            textvariable="CMobile",
            cursor="hand2",
            justify=CENTER,
            relief=SOLID,
        )
        self.CustomerMobileinputField.pack(fill=X, side=LEFT)

        Customer_ntn_gst_Section = Frame(
            CustomerSection, name="customerntnandgstsection"
        )
        Customer_ntn_gst_Section.pack(fill=X, pady=5)

        CustomerNTNLabel = Label(
            Customer_ntn_gst_Section, text="N.T.N:", font=("Sarif", 10, tkFont.BOLD)
        )
        CustomerNTNLabel.pack(fill=X, side=LEFT, ipadx=5)

        self.CustomerNTNinputField = Entry(
            Customer_ntn_gst_Section,
            textvariable="Customerntn",
            cursor="hand2",
            justify=CENTER,
            relief=SOLID,
        )
        self.CustomerNTNinputField.pack(fill=X, side=LEFT, padx=10)

        CustomerGSTLabel = Label(
            Customer_ntn_gst_Section, text="GST:", font=("Sarif", 10, tkFont.BOLD)
        )
        CustomerGSTLabel.pack(fill=X, side=LEFT, padx=5)

        self.CustomerGSTinputField = Entry(
            Customer_ntn_gst_Section,
            textvariable="CGST",
            cursor="hand2",
            justify=CENTER,
            relief=SOLID,
        )
        self.CustomerGSTinputField.pack(fill=X, side=LEFT)

        Customer_Contact_Person_Section = Frame(
            CustomerSection, name="customercontactpersonsection"
        )
        Customer_Contact_Person_Section.pack(fill=X, pady=5)

        Customer_Contact_Person_Label = Label(
            Customer_Contact_Person_Section,
            text="Contact Person:",
            font=("Sarif", 10, tkFont.BOLD),
        )
        Customer_Contact_Person_Label.pack(fill=X, side=LEFT, padx=5)

        self.Customer_Contact_Person_inputField = Entry(
            Customer_Contact_Person_Section,
            textvariable="CCP",
            cursor="hand2",
            justify=CENTER,
            relief=SOLID,
        )
        self.Customer_Contact_Person_inputField.pack(fill=X, side=LEFT)

        CustomerAddressLabel = Label(
            CustomerSection, text="Customer Address:", font=("Sarif", 10, tkFont.BOLD)
        )
        CustomerAddressLabel.pack(fill=X, side=LEFT, ipadx=5)

        self.CustomerAddressinputField = Text(
            CustomerSection,
            width=30,
            height=3,
            cursor="hand2",
            relief=SOLID,
            font=("Sarif", 12, tkFont.NORMAL),
        )
        self.CustomerAddressinputField.pack(fill=X, side=LEFT)

    def detailSection(self):
        DetailSection = LabelFrame(
            self.Section2,
            name="detailsection",
            text="Details",
            fg="black",
            font=("Sarif", 16, tkFont.BOLD),
            padx=15,
            pady=15,
        )
        DetailSection.pack(fill=X)

        DetailTreeHeadings = (
            "SNo",
            "Item Description",
            "Qty",
            "Unit Price",
            "Gross Amount",
            "Discount Per%",
            "Discount Amount",
            "Net Amount",
        )
        DetailTreeItems = []

        DetailTree = ttk.Treeview(
            DetailSection, show="headings", columns=DetailTreeHeadings
        )
        DetailTree.pack(fill=X)

        for col_name in DetailTreeHeadings:
            DetailTree.heading(col_name, text=col_name)

        DetailTree.column(DetailTreeHeadings[0], width=50)  # SNo
        DetailTree.column(DetailTreeHeadings[1], width=350)  # Item Description
        DetailTree.column(DetailTreeHeadings[2], width=100)  # Qty
        DetailTree.column(DetailTreeHeadings[3], width=100)  # Unit Price
        DetailTree.column(DetailTreeHeadings[4], width=100)  # Gross Amount
        DetailTree.column(DetailTreeHeadings[5], width=100)  # Discount Per%
        DetailTree.column(DetailTreeHeadings[6], width=100)  # Discount Amount
        DetailTree.column(DetailTreeHeadings[7], width=100)  # Net Amount

        Add_Detail_Invoice_Generate_Frame = Frame(
            DetailSection, name="adddetailandinvoicegenerateframe"
        )
        Add_Detail_Invoice_Generate_Frame.pack(fill=X)

        def AddDetailFunc():
            AddDetailWin = Tk()
            AddDetailWin.title = "Add New Item"
            AddDetailWin.iconbitmap(TRUE, self.icon_path)

            NewItemFrame = LabelFrame(
                AddDetailWin,
                name="newitemsection",
                text="New Item",
                fg="black",
                font=("Sarif", 16, tkFont.BOLD),
                padx=15,
                pady=15,
            )
            NewItemFrame.pack()

            SNo_Qty_Frame = Frame(NewItemFrame, name="snoandqtyframe")
            SNo_Qty_Frame.pack()

            SNOLabel = LabelFrame(
                SNo_Qty_Frame,
                text="SNo.",
                font=("Sarif", 10, tkFont.BOLD),
                fg="black",
                name="snolabel",
            )
            SNOLabel.pack(fill=X, side=LEFT)

            SNOEntry = Entry(
                SNOLabel,
                textvariable="SNo",
                cursor="hand2",
                justify=CENTER,
                relief=SOLID,
            )
            SNOEntry.pack()

            QtyLabel = LabelFrame(
                SNo_Qty_Frame,
                text="Qty",
                font=("Sarif", 10, tkFont.BOLD),
                fg="black",
                name="qtylabel",
            )
            QtyLabel.pack(fill=X, side=LEFT)

            QtyEntry = Entry(
                QtyLabel,
                textvariable="Qty",
                cursor="hand2",
                justify=CENTER,
                relief=SOLID,
            )
            QtyEntry.pack()

            ItemDescLabel = LabelFrame(
                NewItemFrame,
                text="Item Description",
                font=("Sarif", 10, tkFont.BOLD),
                fg="black",
                name="itemdesclabel",
            )
            ItemDescLabel.pack(fill=X, expand=True)

            ItemDescEntry = Entry(
                ItemDescLabel,
                textvariable="ItemDesc",
                cursor="hand2",
                justify=CENTER,
                relief=SOLID,
            )
            ItemDescEntry.pack(fill=X, expand=True)

            UP_GP_Frame = Frame(NewItemFrame, name="upandgpframe")
            UP_GP_Frame.pack()

            UPriceLabel = LabelFrame(
                UP_GP_Frame,
                text="Unit Price:",
                font=("Sarif", 10, tkFont.BOLD),
                fg="black",
                name="uplabel",
            )
            UPriceLabel.pack(fill=X, side=LEFT)

            UPriceEntry = Entry(
                UPriceLabel,
                textvariable="UPrice",
                cursor="hand2",
                justify=CENTER,
                relief=SOLID,
            )
            UPriceEntry.pack()

            GPriceLabel = LabelFrame(
                UP_GP_Frame,
                text="Gross Price:",
                font=("Sarif", 10, tkFont.BOLD),
                fg="black",
                name="gplabel",
            )
            GPriceLabel.pack(fill=X, side=LEFT)

            GPriceEntry = Entry(
                GPriceLabel,
                textvariable="GPrice",
                cursor="hand2",
                justify=CENTER,
                relief=SOLID,
            )
            GPriceEntry.pack()

            DP_DA_Frame = Frame(NewItemFrame, name="dpanddaframe")
            DP_DA_Frame.pack()

            DiscountPerLabel = LabelFrame(
                DP_DA_Frame,
                text="Discount Per%:",
                font=("Sarif", 10, tkFont.BOLD),
                fg="black",
                name="dplabel",
            )
            DiscountPerLabel.pack(fill=X, side=LEFT)

            DiscountPerEntry = Entry(
                DiscountPerLabel,
                textvariable="DiscountPer",
                cursor="hand2",
                justify=CENTER,
                relief=SOLID,
            )
            DiscountPerEntry.pack()

            DiscountAmountLabel = LabelFrame(
                DP_DA_Frame,
                text="Discount Amount:",
                font=("Sarif", 10, tkFont.BOLD),
                fg="black",
                name="dalabel",
            )
            DiscountAmountLabel.pack(fill=X, side=LEFT)

            DiscountAmountEntry = Entry(
                DiscountAmountLabel,
                textvariable="DiscountAmount",
                cursor="hand2",
                justify=CENTER,
                relief=SOLID,
            )
            DiscountAmountEntry.pack()

            NA_DoneBTN_Frame = Frame(NewItemFrame, name="naanddoneframe")
            NA_DoneBTN_Frame.pack()

            NetAmountLabel = LabelFrame(
                NA_DoneBTN_Frame,
                text="Net Amount:",
                font=("Sarif", 10, tkFont.BOLD),
                fg="black",
                name="nalabel",
            )
            NetAmountLabel.pack(fill=X, side=LEFT)

            NetAmountEntry = Entry(
                NetAmountLabel,
                textvariable="NetAmount",
                cursor="hand2",
                justify=CENTER,
                relief=SOLID,
            )
            NetAmountEntry.pack()

            SNOEntry.delete(0, END)
            QtyEntry.delete(0, END)
            ItemDescEntry.delete(0, END)
            UPriceEntry.delete(0, END)
            GPriceEntry.delete(0, END)
            DiscountPerEntry.delete(0, END)
            DiscountAmountEntry.delete(0, END)
            NetAmountEntry.delete(0, END)

            def AddItemFunc(
                AddDetailWin,
                SNOValue,
                ItemDescValue,
                QtyValue,
                UPriceValue,
                GPriceValue,
                DiscountPerValue,
                DiscountAmountValue,
                NetAmountValue,
            ):
                data = (
                    SNOValue,
                    ItemDescValue,
                    QtyValue,
                    UPriceValue,
                    GPriceValue,
                    DiscountPerValue,
                    DiscountAmountValue,
                    NetAmountValue,
                )
                AddDetailWin.destroy()
                DetailTreeItems.append(data)
                return DetailTree.insert(
                    "",
                    END,
                    values=data,
                )

            AddItemButton = Button(
                NA_DoneBTN_Frame,
                text="Add",
                font=("Sarif", 10, tkFont.BOLD),
                command=lambda: AddItemFunc(
                    AddDetailWin,
                    SNOEntry.get(),
                    ItemDescEntry.get(),
                    QtyEntry.get(),
                    UPriceEntry.get(),
                    GPriceEntry.get(),
                    DiscountPerEntry.get(),
                    DiscountAmountEntry.get(),
                    NetAmountEntry.get(),
                ),
            )
            AddItemButton.pack(fill=X, side=LEFT)

            AddDetailWin.focus()
            AddDetailWin.mainloop()

        AddDetailButton = Button(
            Add_Detail_Invoice_Generate_Frame,
            text="Add Item",
            font=("Sarif", 10, tkFont.BOLD),
            command=AddDetailFunc,
        )
        AddDetailButton.pack(fill=X, side=LEFT, pady=10)

        InvoiceGenerateButton = Button(
            Add_Detail_Invoice_Generate_Frame,
            text="Generate Invoice",
            bg="blue",
            fg="white",
            font=("Sarif", 10, tkFont.BOLD),
            command=lambda: self.makePDF(
                self.CompanyNameinputField.get(),
                self.CompanyLinkinputField.get(),
                self.CompanyAddressinputField.get(1.0, "end-1c"),
                self.SoIDinputField.get(),
                self.SODatePickedVar.get(),
                self.QRefinputField.get(),
                self.DeliveryDatePickedVar.get(),
                self.DescinputField.get(),
                self.payMethodinputField.get(),
                self.PerinputField.get(),
                self.AmountinputField.get(),
                self.EstDatePickedVar.get(),
                self.ActDatePickedVar.get(),
                self.EmpIDinputField.get(),
                self.EmpNameinputField.get(),
                self.EmpDesignationinputField.get(),
                self.EmpDepartmentinputField.get(),
                self.CustomerNameinputField.get(),
                self.CustomerEmailinputField.get(),
                self.CustomerPhoneinputField.get(),
                self.CustomerMobileinputField.get(),
                self.CustomerNTNinputField.get(),
                self.CustomerGSTinputField.get(),
                self.Customer_Contact_Person_inputField.get(),
                self.CustomerAddressinputField.get(1.0, "end-1c"),
                DetailTreeItems,
            ),
        )
        InvoiceGenerateButton.pack(fill=X, side=LEFT, padx=20, pady=10)

    def makePDF(
        self,
        CompanyName_Value,
        CompanyLink_Value,
        CompanyAddress_Value,
        SoID_Value,
        SODatePicked_Value,
        QRef_Value,
        DeliveryDatePicked_Value,
        Desc_Value,
        payMethod_Value,
        Per_Value,
        Amount_Value,
        EstDatePicked_Value,
        ActDatePicked_Value,
        EmpID_Value,
        EmpName_Value,
        EmpDesignation_Value,
        EmpDepartment_Value,
        CustomerName_Value,
        CustomerEmail_Value,
        CustomerPhone_Value,
        CustomerMobile_Value,
        CustomerNTN_Value,
        CustomerGST_Value,
        Customer_Contact_Person__Value,
        CustomerAddress_Value,
        DetailTreeItems,
    ):
        company_info = {
            "name": CompanyName_Value,
            "address": CompanyAddress_Value,
            "phone": CustomerPhone_Value,
            "website": CompanyLink_Value,
        }
        customer_info = {
            "Name": CustomerName_Value,
            "Email": CustomerEmail_Value,
            "Phone": CustomerPhone_Value,
            "Mobile": CustomerMobile_Value,
            "NTN": CustomerNTN_Value,
            "GST": CustomerGST_Value,
            "Contact Person": Customer_Contact_Person__Value,
            "Address": CustomerAddress_Value,
        }
        sale_order_info = {
            "Sales Order ID": SoID_Value,
            "Sales Order Date": SODatePicked_Value,
            "Quotation Ref": QRef_Value,
            "Delivery Date": DeliveryDatePicked_Value,
        }
        sales_person = [
            [EmpID_Value, EmpName_Value, EmpDesignation_Value, EmpDepartment_Value]
        ]
        payment_schedule = [
            [
                Desc_Value,
                EstDatePicked_Value,
                ActDatePicked_Value,
                Per_Value,
                Amount_Value,
                payMethod_Value,
            ]
        ]
        details = DetailTreeItems

        new_pdf = PDF()
        new_pdf.generate(
            company_info,
            customer_info,
            sale_order_info,
            sales_person,
            payment_schedule,
            details,
        )

    def mainFooter(self, mainFrame):
        footerFrame = Frame(mainFrame, name="footerframe")
        footerFrame.pack(fill=X, side=BOTTOM)

        developerTag = Label(
            footerFrame,
            text="Developed By AL-FEROZ TECH",
            fg="gray",
            font=("Sarif", 8, tkFont.ITALIC),
        )
        developerTag.pack(fill=X, side=RIGHT)


if __name__ == "__main__":
    root = Tk()
    app = InvoiceGenerator(root)
    root.mainloop()

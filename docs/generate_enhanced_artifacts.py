from pathlib import Path
import subprocess

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from pptx import Presentation
from pptx.dml.color import RGBColor as PPTRGB
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches as PInches, Pt as PPt


ROOT = Path(__file__).resolve().parent
UML = ROOT / "uml"
REPORT = ROOT / "report"
PRESENTATION = ROOT / "presentation"
REPORT.mkdir(exist_ok=True)
PRESENTATION.mkdir(exist_ok=True)

NAVY = "10233F"
TEAL = "13B8A6"
CYAN = "53C7D9"
BLUE = "3A74D8"
PALE = "EDF5F7"
WHITE = "FFFFFF"
INK = "203047"
GOLD = "F2B84B"
GREEN = "57B894"
RED = "DE6B6B"


DOT = {
    "use-case": r'''digraph G { rankdir=LR; node [shape=ellipse, style=filled, fillcolor="#E8F1FF", fontname="Arial"]; Guest[label="Guest"]; User[label="User"]; Admin[label="Admin"]; subgraph cluster_app { label="ParQHub web application"; color="#9BB7D4"; Browse[label="Browse home/login"]; Register[label="Register and vehicles"]; Login[label="Sign in"]; Search[label="Search parking"]; Book[label="Book slot"]; EntryExit[label="Entry and exit"]; Pay[label="Process payment"]; Notify[label="Notifications"]; Dashboard[label="Role dashboard"]; Manage[label="Manage data"]; Reports[label="Generate reports"]; } Guest->Browse; Guest->Register; Guest->Login; User->Search; User->Book; User->EntryExit; User->Pay; User->Notify; Admin->Dashboard; Admin->Manage; Admin->Reports; }''',
    "class-diagram": r'''digraph G { rankdir=LR; node [shape=record, fontname="Arial", fontsize=10]; User[label="{User|id; name; email; password; contactNo}"]; Admin[label="{Admin|id; name; email; password; role}"]; Vehicle[label="{Vehicle|id; vehicleNo; vehicleType; brand; model; color}"]; Lot[label="{ParkingLot|id; city; location; totalSlots; priceHr}"]; Slot[label="{ParkingSlot|id; status}"]; Booking[label="{Booking|id; startTime; endTime; paymentStatus}"]; Payment[label="{Payment|id; amount; method; status; timestamp}"]; Log[label="{VehicleLog|id; entryTime; exitTime}"]; Notification[label="{Notification|id; type; description; timestamp}"]; Report[label="{Report|id; type; generatedDate}"]; User->Vehicle[label="1..*"]; User->Booking[label="1..*"]; User->Notification[label="1..*"]; Lot->Slot[label="1..*"]; Booking->User[label="*..1"]; Booking->Vehicle[label="*..1"]; Booking->Slot[label="*..1"]; Payment->Booking[label="*..1"]; Log->Vehicle[label="*..1"]; Log->Lot[label="*..1"]; Notification->User[label="*..0..1"]; Notification->Admin[label="*..0..1"]; Report->Admin[label="*..1"]; }''',
    "activity": r'''digraph G { rankdir=TB; node [shape=box, style="rounded,filled", fillcolor="#F4F8FC", fontname="Arial"]; start[shape=circle,label="",width=.2]; login[shape=diamond,label="Logged in?"]; search[label="Search city/location"]; slots[label="Load available slots"]; select[label="Select vehicle and slot"]; available[shape=diamond,label="Slot available?"]; book[label="Create pending booking"]; entry[label="Record entry; slot OCCUPIED"]; exitStep[label="Record exit"]; duration[label="Calculate duration"]; pay[label="Process payment"]; done[label="Complete booking; slot AVAILABLE"]; end[shape=doublecircle,label="",width=.2]; start->login; login->search[label="yes"]; login->login[label="no"]; search->slots->select->available; available->search[label="no"]; available->book[label="yes"]; book->entry->exitStep->duration->pay->done->end; }''',
    "state-machine": r'''digraph G { rankdir=LR; node [shape=ellipse, style=filled, fillcolor="#FFF8E6", fontname="Arial"]; start[shape=point]; available[label="AVAILABLE"]; booked[label="BOOKED"]; occupied[label="OCCUPIED"]; start->available; available->booked[label="booking created"]; booked->occupied[label="vehicle entry"]; occupied->available[label="vehicle exit + payment"]; booked->available[label="cancellation flow"]; pending[label="Payment PENDING"]; completed[label="Payment COMPLETED"]; pending->completed[label="strategy returns success"]; }''',
    "component": r'''digraph G { rankdir=LR; node [shape=box, style="rounded,filled", fillcolor="#EAF7EE", fontname="Arial"]; browser[label="Browser"]; mvc[label="Spring MVC controllers"]; views[label="Thymeleaf templates"]; services[label="Service layer"]; strategies[label="Strategy contexts"]; repos[label="Spring Data JPA repositories"]; db[shape=cylinder,label="MySQL parqhub_db"]; mail[label="JavaMailSender"]; browser->mvc; mvc->views; mvc->services; mvc->repos; services->strategies; services->repos; mvc->mail; repos->db; }''',
    "deployment": r'''digraph G { rankdir=LR; node [shape=box3d, style=filled, fillcolor="#EEF2FF", fontname="Arial"]; client[label="Browser"]; app[label="JDK 21 host\nSpring Boot 3.3.4\nPort 8080"]; db[shape=cylinder,label="MySQL parqhub_db"]; smtp[label="SMTP service"]; client->app[label="HTTP"]; app->db[label="JDBC"]; app->smtp[label="SMTP"]; }''',
    "package": r'''digraph G { rankdir=TB; node [shape=folder, style=filled, fillcolor="#F4F4F4", fontname="Arial"]; root[label="com.prmplatform.parqhub"]; controller[label="controller"]; admin[label="controller.admin"]; user[label="controller.user"]; model[label="model"]; repo[label="repository"]; service[label="service"]; strategy[label="strategy"]; payment[label="strategy.payment"]; auth[label="strategy.adminauth"]; root->controller; root->model; root->repo; root->service; root->strategy; controller->admin; controller->user; strategy->payment; strategy->auth; controller->service; controller->repo; service->repo; service->strategy; repo->model; }''',
    "entity-relationship": r'''digraph G { rankdir=LR; node [shape=record, fontname="Arial", fontsize=9]; Users[label="{USERS|user_id PK|name|email UK|password|contact_no}"]; Admins[label="{ADMINS|admin_id PK|name|email UK|password|role}"]; Vehicle[label="{VEHICLE|vehicle_id PK|user_id FK|vehicle_no UK|vehicle_type}"]; Lot[label="{PARKINGLOT|lot_id PK|city|location|total_slots|price_hr}"]; Slot[label="{PARKINGSLOT|slot_id PK|lot_id FK|status}"]; Booking[label="{BOOKING|booking_id PK|user_id FK|vehicle_id FK|slot_id FK|start_time|end_time|payment_status}"]; Payment[label="{PAYMENT|payment_id PK|booking_id FK|amount|method|status|timestamp}"]; Log[label="{VEHICLELOG|log_id PK|vehicle_id FK|lot_id FK|entry_time|exit_time}"]; Report[label="{REPORT|report_id PK|admin_id FK|type|generated_date}"]; Notification[label="{NOTIFICATION|notification_id PK|user_id FK|admin_id FK|type|description|timestamp}"]; Users->Vehicle; Users->Booking; Users->Notification; Admins->Report; Admins->Notification; Lot->Slot; Vehicle->Booking; Slot->Booking; Booking->Payment; Vehicle->Log; Lot->Log; }''',
    "sequence-booking": r'''digraph G { rankdir=LR; node [shape=box, style=filled, fillcolor="#F4F8FC", fontname="Arial"]; U[label="User"]; UI[label="Thymeleaf UI"]; C[label="FindParkingController"]; V[label="VehicleRepository"]; S[label="ParkingSlotRepository"]; B[label="BookingRepository"]; D[label="MySQL"]; U->UI->C->V; C->S; C->B; V->D; S->D; B->D; }''',
    "sequence-payment": r'''digraph G { rankdir=LR; node [shape=box, style=filled, fillcolor="#F4F8FC", fontname="Arial"]; U[label="User"]; C[label="PaymentController"]; P[label="PaymentProcessingService"]; X[label="PaymentContext"]; T[label="Card/Cash/Arrival"]; R[label="Repositories"]; M[label="JavaMailSender"]; U->C->P->X->T; C->R; C->M; }''',
    "sequence-entry-exit": r'''digraph G { rankdir=LR; node [shape=box, style=filled, fillcolor="#F4F8FC", fontname="Arial"]; U[label="User"]; C[label="VehicleLogController"]; B[label="BookingRepository"]; L[label="VehicleLogRepository"]; S[label="ParkingSlotRepository"]; U->C->B; C->L; C->S; }''',
}


def export_diagrams():
    for name, source in DOT.items():
        dot_file = UML / f".{name}.dot"
        dot_file.write_text(source)
        for fmt in ("svg", "png"):
            subprocess.run(["dot", f"-T{fmt}", str(dot_file), "-o", str(UML / f"{name}.{fmt}")], check=True)
        dot_file.unlink()


def set_cell_shading(cell, color):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), color)
    tc_pr.append(shd)


def set_repeat_table_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def add_doc_heading(doc, text, level=1):
    p = doc.add_heading(text, level=level)
    p.runs[0].font.color.rgb = RGBColor(16, 35, 63)
    return p


def add_doc_table(doc, headers, rows):
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = header
        set_cell_shading(cell, NAVY)
        for run in cell.paragraphs[0].runs:
            run.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
    set_repeat_table_header(table.rows[0])
    for ridx, row in enumerate(rows):
        cells = table.add_row().cells
        for i, value in enumerate(row):
            cells[i].text = str(value)
            cells[i].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            if ridx % 2:
                set_cell_shading(cells[i], PALE)
    return table


def add_page_number(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = paragraph.add_run("Page ")
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    paragraph._p.append(fld)


def generate_report():
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Inches(0.65)
    sec.bottom_margin = Inches(0.65)
    sec.left_margin = Inches(0.75)
    sec.right_margin = Inches(0.75)

    styles = doc.styles
    styles["Normal"].font.name = "Aptos"
    styles["Normal"].font.size = Pt(10)
    styles["Title"].font.name = "Aptos Display"
    styles["Title"].font.color.rgb = RGBColor(16, 35, 63)
    for level, size in ((1, 20), (2, 14), (3, 11)):
        styles[f"Heading {level}"].font.name = "Aptos Display"
        styles[f"Heading {level}"].font.size = Pt(size)
        styles[f"Heading {level}"].font.color.rgb = RGBColor(19, 184, 166)

    footer = sec.footer.paragraphs[0]
    footer.text = "ParQHub | Object-Oriented Modelling Report"
    add_page_number(sec.footer.add_paragraph())

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.add_run("ParQHub").bold = True
    title.runs[0].font.size = Pt(38)
    title.runs[0].font.color.rgb = RGBColor(16, 35, 63)

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = sub.add_run("Object-Oriented Modelling\n& OOPS Assignment Report")
    run.bold = True
    run.font.size = Pt(23)
    run.font.color.rgb = RGBColor(19, 184, 166)

    line = doc.add_paragraph()
    line.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rr = line.add_run("SMART PARKING MANAGEMENT SYSTEM")
    rr.bold = True
    rr.font.size = Pt(11)
    rr.font.color.rgb = RGBColor(58, 116, 216)

    doc.add_paragraph()
    team = doc.add_table(rows=4, cols=2)
    team.alignment = WD_TABLE_ALIGNMENT.CENTER
    team.style = "Table Grid"
    entries = [("Submitted by", "Paarth Arora"), ("Team Member", "Mannat Jain"), ("Team Member", "Kajal Jayaswal"), ("Technology", "Java 21 / Spring Boot 3.3.4")]
    for i, (label, value) in enumerate(entries):
        team.rows[i].cells[0].text = label
        team.rows[i].cells[1].text = value
        set_cell_shading(team.rows[i].cells[0], NAVY)
        for run in team.rows[i].cells[0].paragraphs[0].runs:
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.bold = True
    doc.add_page_break()

    add_doc_heading(doc, "Document Map", 1)
    add_doc_table(doc, ["Section", "Focus"], [
        ("1", "Executive overview and problem statement"),
        ("2", "Stakeholders, scope and requirements"),
        ("3", "Logical architecture and runtime deployment"),
        ("4", "Domain model and JPA relationships"),
        ("5", "Object-oriented concepts with code evidence"),
        ("6", "Design patterns and role-based behavior"),
        ("7", "SOLID principles and trade-off analysis"),
        ("8", "UML behavioural and structural models"),
        ("9", "Security, quality, testing and limitations"),
        ("10", "Conclusion and future enhancements"),
    ])

    add_doc_heading(doc, "1. Executive Overview", 1)
    doc.add_paragraph(
        "ParQHub is a layered web application that manages the complete lifecycle of a parking transaction: "
        "user registration, vehicle registration, parking discovery, slot booking, vehicle entry, occupancy, "
        "vehicle exit, payment, reporting and notification. It replaces disconnected manual activities with "
        "a Spring MVC application backed by JPA repositories and MySQL."
    )
    add_doc_heading(doc, "1.1 Problem Statement", 2)
    doc.add_paragraph(
        "The system must prevent double allocation of a limited parking slot, represent temporary and persistent "
        "states consistently, calculate a payable duration, support several payment methods, and provide role-specific "
        "administrative visibility. The model must also remain understandable and extensible as the application grows."
    )
    add_doc_heading(doc, "1.2 Project Objectives", 2)
    for item in (
        "Centralize parking discovery and reservation workflows.",
        "Represent users, vehicles, lots, slots, bookings, payments and logs as connected domain objects.",
        "Keep controllers thin and delegate business behavior to services and strategy implementations.",
        "Demonstrate encapsulation, abstraction, polymorphism and composition in executable Java code.",
        "Visualize system structure and behavior using UML."
    ):
        doc.add_paragraph(item, style="List Bullet")

    doc.add_page_break()
    add_doc_heading(doc, "2. Requirements and Stakeholders", 1)
    add_doc_heading(doc, "2.1 Stakeholder Matrix", 2)
    add_doc_table(doc, ["Stakeholder", "Goals", "Primary interactions"], [
        ("Guest", "Discover and access the system", "Home, registration, login"),
        ("User / Driver", "Find, reserve and pay for parking", "Vehicle, slot, booking, log, payment"),
        ("Operations Manager", "Monitor and manage parking operations", "Occupancy and slot management"),
        ("Finance Officer", "Review revenue and pending payments", "Financial dashboards and reports"),
        ("Customer Service", "Support users and bookings", "Customer-oriented dashboard"),
        ("Security Supervisor", "Monitor active vehicles and incidents", "Vehicle logs and notifications"),
        ("IT Support", "Maintain and observe system health", "Health and notification views"),
        ("Super Admin", "Access all system functions", "All role sections"),
    ])
    add_doc_heading(doc, "2.2 Functional Requirements", 2)
    for item in (
        "Register a user and authenticate it using an HTTP session.",
        "Register, update and view multiple vehicles associated with a user.",
        "Search parking lots by city, location, price and availability.",
        "Create a booking only from an available parking slot.",
        "Record vehicle entry and exit, update occupancy and calculate parking duration.",
        "Process card, cash or arrival payment methods and mark the booking completed.",
        "Send a payment confirmation through JavaMailSender when mail configuration is available.",
        "Allow an administrator to view role-filtered operational and financial information.",
    ):
        doc.add_paragraph(item, style="List Bullet")
    add_doc_heading(doc, "2.3 Non-Functional Requirements", 2)
    add_doc_table(doc, ["Quality", "Implementation evidence"], [
        ("Usability", "Thymeleaf pages for separate user and administrator journeys"),
        ("Maintainability", "Controllers, services, repositories and strategy packages are separated"),
        ("Reliability", "JPA persistence and explicit payment and booking status values"),
        ("Extensibility", "Payment and authentication strategies implement focused interfaces"),
        ("Portability", "Spring Boot, Maven and JDBC provide a conventional deployment stack"),
        ("Security", "Session checks and role-based authentication service guard protected views"),
    ])

    doc.add_page_break()
    add_doc_heading(doc, "3. Architecture and Technology", 1)
    add_doc_heading(doc, "3.1 Layered Architecture", 2)
    add_doc_table(doc, ["Layer", "Representative code", "Responsibility"], [
        ("Presentation", "src/main/resources/templates", "Thymeleaf HTML views"),
        ("Web", "controller and controller.* packages", "HTTP input, session checks and view selection"),
        ("Application", "service package", "Parking, payment and role-authentication behavior"),
        ("Strategy", "strategy.payment / strategy.adminauth", "Interchangeable payment and role behavior"),
        ("Persistence", "repository package", "Spring Data JPA data access"),
        ("Domain", "model package", "JPA entities and DTOs"),
    ])
    add_doc_heading(doc, "3.2 Technology Stack", 2)
    add_doc_table(doc, ["Component", "Version / choice"], [
        ("Language", "Java 21"),
        ("Framework", "Spring Boot 3.3.4"),
        ("Web", "Spring MVC and Thymeleaf"),
        ("Persistence", "Spring Data JPA / Hibernate"),
        ("Database", "MySQL; H2 dependency available for testing"),
        ("Mail", "Spring Mail / JavaMailSender"),
        ("Build", "Maven"),
        ("Diagrams", "Graphviz / Mermaid exports"),
    ])
    add_doc_heading(doc, "3.3 Runtime Flow", 2)
    doc.add_paragraph(
        "The browser submits HTTP requests to controllers. Controllers check HTTP session state, call repositories "
        "or services, and return a Thymeleaf view or redirect. Services and strategy contexts contain reusable behavior "
        "rather than HTML concerns. Repositories translate object operations into JPA persistence against MySQL."
    )
    for name in ("component", "deployment"):
        doc.add_picture(str(UML / f"{name}.png"), width=Inches(6.15))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(name.replace("-", " ").title(), style="Caption")

    doc.add_page_break()
    add_doc_heading(doc, "4. Object Model", 1)
    add_doc_heading(doc, "4.1 Domain Classes", 2)
    add_doc_table(doc, ["Class", "Key state", "Responsibility"], [
        ("User", "identity, credentials, contact details", "Owns vehicles, bookings and notifications"),
        ("Vehicle", "number, type, brand, model, color", "Identifies a vehicle in logs and bookings"),
        ("ParkingLot", "city, location, capacity and hourly price", "Contains a collection of parking slots"),
        ("ParkingSlot", "lot and slot status", "Represents AVAILABLE, BOOKED or OCCUPIED capacity"),
        ("Booking", "user, vehicle, slot, times and payment status", "Connects one parking transaction"),
        ("Payment", "amount, method, status and timestamp", "Records a completed or pending transaction"),
        ("VehicleLog", "entry and exit times", "Measures physical parking occupancy"),
        ("Admin", "identity and role", "Receives role-specific dashboard access"),
        ("Notification", "type and description", "Carries operational or security messages"),
        ("Report", "type and generated date", "Represents an administrative report"),
    ])
    add_doc_heading(doc, "4.2 Association Analysis", 2)
    doc.add_paragraph(
        "The model uses composition and association rather than inheritance-heavy domain design. "
        "A ParkingLot aggregates ParkingSlot objects. A Booking associates a User, Vehicle and ParkingSlot. "
        "A Payment belongs to a Booking. VehicleLog associates a Vehicle with a ParkingLot. "
        "Notification may associate a User or an Admin, and Report is associated with an Admin."
    )
    doc.add_picture(str(UML / "class-diagram.png"), width=Inches(6.25))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("Class Diagram", style="Caption")

    doc.add_page_break()
    add_doc_heading(doc, "5. Object-Oriented Principles", 1)
    principles = [
        ("Encapsulation", "ParkingSlot, Booking and Payment keep state private and expose accessors.", "Private JPA fields plus public getters/setters."),
        ("Abstraction", "PaymentStrategy and AdminAuthenticationStrategy hide concrete behavior behind interfaces.", "Strategy implementations are selected by context objects."),
        ("Polymorphism", "PaymentContext uses one interface with Card, Cash and Arrival implementations.", "AdminAuthenticationContext similarly delegates to role strategies."),
        ("Composition", "ParkingLot contains slots; Booking references user, vehicle and slot.", "Runtime behaviour emerges from object collaboration."),
        ("Inheritance", "It is intentionally limited in the domain model.", "Interfaces and composition avoid an unnecessary domain hierarchy."),
    ]
    add_doc_table(doc, ["Principle", "Evidence", "Effect"], principles)
    add_doc_heading(doc, "5.1 Encapsulation Example", 2)
    doc.add_paragraph(
        "Booking exposes paymentStatus through getPaymentStatus and setPaymentStatus instead of exposing its field directly. "
        "LocalDateTime endTime and slot references remain private JPA state."
    )
    add_doc_heading(doc, "5.2 Polymorphism Example", 2)
    doc.add_paragraph(
        "PaymentContext stores a PaymentStrategy. processPayment, getPaymentMethodName and isValidPaymentMethod delegate "
        "to the selected concrete strategy. Switching from CardPaymentStrategy to CashPaymentStrategy changes behaviour "
        "without changing controller code."
    )
    add_doc_heading(doc, "5.3 Composition Example", 2)
    doc.add_paragraph(
        "ParkingSlot holds a reference to a ParkingLot, and Booking holds references to the user, vehicle and slot. "
        "These relationships match the transaction and allow JPA to persist and load a navigable object graph."
    )

    doc.add_page_break()
    add_doc_heading(doc, "6. Design Patterns", 1)
    add_doc_table(doc, ["Pattern", "Evidence", "Benefit"], [
        ("MVC", "Controllers, entities and Thymeleaf templates", "Separates request handling, domain state and presentation"),
        ("Repository", "Spring Data JpaRepository interfaces", "Removes persistence boilerplate and exposes focused queries"),
        ("Service Layer", "PaymentProcessingService, ParkingLotService", "Keeps business logic out of controllers"),
        ("Strategy", "PaymentContext / AdminAuthenticationContext", "Changes behavior through interchangeable implementations"),
        ("DTO", "ParkingSlotDTO / VehicleDTO", "Transfers a focused data shape instead of exposing entities"),
        ("Dependency Injection", "Spring constructor and field injection", "Decouples components and enables framework-managed composition"),
    ])
    add_doc_heading(doc, "6.1 Payment Strategy Sequence", 2)
    doc.add_paragraph(
        "PaymentController calculates the parking charge, validates the requested method and asks "
        "PaymentProcessingService to process the payment. The service selects the matching strategy. "
        "The strategy creates a Payment object with amount, method, completed status and timestamp."
    )
    doc.add_picture(str(UML / "sequence-payment.png"), width=Inches(6.25))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    add_doc_heading(doc, "6.2 Administrator Strategy Sequence", 2)
    doc.add_paragraph(
        "AdminAuthenticationService supplies role-specific behavior through AdminAuthenticationContext. "
        "The selected strategy validates a role and the service uses role access checks to expose dashboard sections "
        "for operations, finance, customer service, security or IT."
    )

    doc.add_page_break()
    add_doc_heading(doc, "7. SOLID Assessment", 1)
    add_doc_table(doc, ["Principle", "Evidence", "Caveat"], [
        ("Single Responsibility", "Controllers handle web concerns; services and repositories separate behavior", "Some controllers contain substantial workflow and email logic"),
        ("Open/Closed", "New payment or authentication strategies can be added behind interfaces", "Concrete strategy registration must remain complete"),
        ("Liskov Substitution", "Concrete strategies satisfy their context-facing interfaces", "No alternate implementation can violate the interface contract"),
        ("Interface Segregation", "Payment and authentication interfaces contain only behavior needed by their contexts", "Repository interfaces are framework-generated and broader"),
        ("Dependency Inversion", "Contexts depend on strategy interfaces; Spring injects implementations", "Controllers still reference concrete repositories and services"),
    ])
    add_doc_heading(doc, "7.1 Trade-Off Analysis", 2)
    doc.add_paragraph(
        "The architecture is intentionally practical for a single Spring MVC application. "
        "Direct repository use is easier to teach and integrate than a separate application-service API, "
        "but it also means some controllers have more orchestration responsibility. "
        "The strategy pattern improves the payment and role-authentication areas without forcing a full framework abstraction."
    )

    doc.add_page_break()
    add_doc_heading(doc, "8. UML and Behavioral Models", 1)
    uml_sections = [
        ("Use Case", "use-case", "Shows guest, user and administrator goals."),
        ("Class", "class-diagram", "Shows domain classes, fields and object associations."),
        ("Entity Relationship", "entity-relationship", "Shows database tables and foreign-key relationships."),
        ("Booking Sequence", "sequence-booking", "Shows search, slot selection and booking persistence."),
        ("Payment Sequence", "sequence-payment", "Shows charge calculation and strategy execution."),
        ("Entry and Exit Sequence", "sequence-entry-exit", "Shows occupancy transitions and log persistence."),
        ("Activity", "activity", "Shows decisions and actions through the parking journey."),
        ("State Machine", "state-machine", "Shows slot, booking and payment state transitions."),
        ("Package", "package", "Shows source-package separation."),
    ]
    for title, name, description in uml_sections:
        add_doc_heading(doc, f"8.{uml_sections.index((title, name, description)) + 1} {title}", 2)
        doc.add_paragraph(description)
        doc.add_picture(str(UML / f"{name}.png"), width=Inches(6.25))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_page_break()

    add_doc_heading(doc, "9. Operational States and Concurrency", 1)
    add_doc_heading(doc, "9.1 Parking Slot Lifecycle", 2)
    add_doc_table(doc, ["State", "Meaning", "Transition"], [
        ("AVAILABLE", "Slot is not allocated", "A booking changes it to BOOKED"),
        ("BOOKED", "A reservation owns the slot", "Entry changes it to OCCUPIED"),
        ("OCCUPIED", "Vehicle is physically present", "Exit and completion return it to AVAILABLE"),
    ])
    add_doc_heading(doc, "9.2 Booking and Payment Lifecycle", 2)
    doc.add_paragraph(
        "Booking.paymentStatus begins Pending. The payment process creates a Payment, marks the Booking Completed "
        "and returns the slot to AVAILABLE. Failed remains a possible enum value, but the presented workflow is "
        "designed to report processing errors rather than silently change state."
    )
    add_doc_heading(doc, "9.3 Concurrency Consideration", 2)
    doc.add_paragraph(
        "A robust production implementation should guarantee that two concurrent requests cannot reserve the final "
        "AVAILABLE slot. The current code delegates slot allocation to repository queries and controller/service logic "
        "and does not establish an explicit database locking strategy. A future version should use a transaction with "
        "pessimistic locking or a conditional state update."
    )

    doc.add_page_break()
    add_doc_heading(doc, "10. Quality, Testing and Limitations", 1)
    add_doc_heading(doc, "10.1 Testing Evidence", 2)
    doc.add_paragraph(
        "The repository contains an application-context test and the Maven build output includes Surefire reports. "
        "The test suite is small and should be expanded to cover strategy selection, slot transitions, payment "
        "calculation and unauthorized session access."
    )
    add_doc_heading(doc, "10.2 Recommended Test Cases", 2)
    add_doc_table(doc, ["Area", "Test case", "Expected result"], [
        ("Booking", "No available slot", "Booking is rejected"),
        ("Booking", "Available slot", "Status becomes BOOKED"),
        ("Entry", "Matching vehicle and booking", "Status becomes OCCUPIED"),
        ("Exit", "Completed exit", "Log exitTime is set"),
        ("Payment", "Card / Cash / Arrival", "Correct strategy returns Completed payment"),
        ("Authorization", "User or admin session missing", "Protected view redirects to login"),
        ("Admin auth", "Role strategy mismatch", "Authentication fails"),
    ])
    add_doc_heading(doc, "10.3 Current Limitations", 2)
    for item in (
        "Dashboard system uptime is represented by a placeholder.",
        "Mail delivery failures are logged but do not fail or roll back the payment workflow.",
        "Password and session security should be hardened before public deployment.",
        "Explicit pessimistic or optimistic concurrency control is not demonstrated.",
        "Automated integration and role-matrix tests are not comprehensive.",
    ):
        doc.add_paragraph(item, style="List Bullet")

    doc.add_page_break()
    add_doc_heading(doc, "11. Conclusion", 1)
    doc.add_paragraph(
        "ParQHub demonstrates a coherent object-oriented parking-management solution. "
        "The Spring MVC structure separates concerns, JPA entities represent the domain, repositories provide "
        "persistence access, services isolate orchestration, and strategy contexts make payment and administrator "
        "authentication behavior replaceable."
    )
    doc.add_paragraph(
        "The UML models are based on executable source packages and SQL relationships rather than an unrelated "
        "fictional design. The report identifies both strong architectural evidence and current limitations, "
        "making it suitable for academic evaluation and implementation discussion."
    )
    add_doc_heading(doc, "11.1 Future Enhancements", 2)
    for item in (
        "Add Spring Security, password hashing and CSRF protection.",
        "Use database locking and transactional booking/payment workflows.",
        "Add REST contracts and comprehensive API/integration testing.",
        "Replace placeholders with real telemetry and monitoring.",
        "Support reservations with start and end times in the core booking flow.",
    ):
        doc.add_paragraph(item, style="List Bullet")

    doc.save(REPORT / "ParQHub_OOM_Report.docx")


def rgb(hex_value):
    return PPTRGB(int(hex_value[0:2], 16), int(hex_value[2:4], 16), int(hex_value[4:6], 16))


def set_shape_fill(shape, color):
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(color)


def set_shape_line(shape, color=None, width=1):
    if color is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = rgb(color)
        shape.line.width = PInches(width / 72)


def add_text(slide, text, left, top, width, height, size=20, color=INK,
             bold=False, align=PP_ALIGN.LEFT, font="Aptos", margin=0.08,
             valign=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(PInches(left), PInches(top),
                                   PInches(width), PInches(height))
    frame = box.text_frame
    frame.clear()
    frame.margin_left = PInches(margin)
    frame.margin_right = PInches(margin)
    frame.margin_top = PInches(margin)
    frame.margin_bottom = PInches(margin)
    frame.word_wrap = True
    frame.vertical_anchor = valign
    para = frame.paragraphs[0]
    para.text = text
    para.alignment = align
    para.font.name = font
    para.font.size = PPt(size)
    para.font.bold = bold
    para.font.color.rgb = rgb(color)
    return box


def add_rich_text(slide, items, left, top, width, height,
                  size=18, color=INK, bullet=True):
    box = slide.shapes.add_textbox(PInches(left), PInches(top),
                                   PInches(width), PInches(height))
    frame = box.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.margin_left = PInches(0.12)
    frame.margin_right = PInches(0.08)
    for i, item in enumerate(items):
        para = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
        para.text = item
        para.font.name = "Aptos"
        para.font.size = PPt(size)
        para.font.color.rgb = rgb(color)
        para.space_after = PPt(12)
        if bullet:
            para.text = "• " + item
    return box


def add_header(slide, section, title, number):
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, PInches(0), PInches(0),
                                 PInches(13.333), PInches(0.12))
    set_shape_fill(bar, TEAL)
    set_shape_line(bar)
    add_text(slide, section.upper(), 0.65, 0.28, 2.7, 0.32, 11, TEAL, True)
    add_text(slide, title, 0.65, 0.62, 11.7, 0.58, 29, NAVY, True)
    add_text(slide, f"{number:02}", 12.05, 0.55, 0.65, 0.45, 16, BLUE, True,
             align=PP_ALIGN.RIGHT)
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, PInches(0.65),
                                  PInches(1.22), PInches(12.03), PInches(0.018))
    set_shape_fill(line, "D8E4EA")
    set_shape_line(line)


def add_footer(slide, source="PARQHUB • OOPS ASSIGNMENT"):
    add_text(slide, source, 0.65, 7.12, 8.5, 0.22, 8, "7A8D9E", True)
    add_text(slide, "TEAM: PAARTH • MANNAT • KAJAL", 8.15, 7.12, 4.52, 0.22,
             8, "7A8D9E", True, align=PP_ALIGN.RIGHT)


def add_metric_card(slide, left, top, width, label, value, accent=TEAL):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                  PInches(left), PInches(top),
                                  PInches(width), PInches(1.25))
    set_shape_fill(card, WHITE)
    set_shape_line(card, "D6E2E8", 1)
    strip = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                   PInches(left), PInches(top),
                                   PInches(0.08), PInches(1.25))
    set_shape_fill(strip, accent)
    set_shape_line(strip)
    add_text(slide, value, left + 0.22, top + 0.15, width - 0.35, 0.55,
             24, NAVY, True)
    add_text(slide, label, left + 0.22, top + 0.74, width - 0.35, 0.28,
             10, "6D8192", True)


def add_diagram_slide(prs, number, section, title, name, caption):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(WHITE)
    add_header(slide, section, title, number)
    panel = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                   PInches(0.65), PInches(1.48),
                                   PInches(12.02), PInches(5.42))
    set_shape_fill(panel, "F7FAFB")
    set_shape_line(panel, "D7E4EA", 1)
    image = UML / f"{name}.png"
    if image.exists():
        slide.shapes.add_picture(str(image), PInches(0.9), PInches(1.65),
                                 width=PInches(11.52), height=PInches(4.78))
    add_text(slide, caption, 1.05, 6.38, 11.22, 0.32, 10, "668092",
             True, align=PP_ALIGN.CENTER)
    add_footer(slide)
    return slide


def generate_presentation():
    prs = Presentation()
    prs.slide_width = PInches(13.333)
    prs.slide_height = PInches(7.5)

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(NAVY)
    accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, PInches(0),
                                    PInches(0), PInches(0.16),
                                    PInches(7.5))
    set_shape_fill(accent, TEAL)
    set_shape_line(accent)
    add_text(slide, "PARQHUB", 0.85, 0.72, 5.2, 0.48, 16, CYAN, True)
    add_text(slide, "Object-Oriented\nModelling", 0.85, 1.45, 7.3, 1.55,
             36, WHITE, True)
    add_text(slide, "A visual journey through architecture, UML and design patterns",
             0.85, 3.18, 7.25, 0.66, 18, "C9D7E3")
    add_text(slide, "Paarth Arora  •  Mannat Jain  •  Kajal Jayaswal",
             0.85, 6.55, 7.3, 0.38, 12, WHITE, True)
    circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, PInches(9.05),
                                    PInches(1.15), PInches(3.25), PInches(3.25))
    set_shape_fill(circle, "183C63")
    set_shape_line(circle, CYAN, 2)
    add_text(slide, "P", 9.05, 1.72, 3.25, 1.15, 54, CYAN, True,
             align=PP_ALIGN.CENTER)
    add_text(slide, "SMART PARKING", 9.05, 3.12, 3.25, 0.35, 13,
             WHITE, True, align=PP_ALIGN.CENTER)
    add_text(slide, "OOM + OOPS", 9.05, 5.42, 3.25, 0.48, 17,
             GOLD, True, align=PP_ALIGN.CENTER)

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(WHITE)
    add_header(slide, "Overview", "Parking, connected as objects", 2)
    add_text(slide,
             "One system links discovery, capacity, reservations, physical occupancy and payment.",
             0.65, 1.48, 7.2, 0.82, 20, INK, True)
    add_rich_text(slide, [
        "Search city, location and hourly price",
        "Register vehicles and create bookings",
        "Track entry, exit and slot state",
        "Calculate duration and process payment",
        "Serve role-specific administrator views",
    ], 0.65, 2.42, 5.15, 3.55, 17)
    add_metric_card(slide, 6.25, 1.62, 2.0, "CORE LANGUAGE", "Java 21", TEAL)
    add_metric_card(slide, 8.55, 1.62, 2.0, "FRAMEWORK", "Spring Boot", BLUE)
    add_metric_card(slide, 10.85, 1.62, 1.82, "PERSISTENCE", "JPA + SQL", GOLD)
    add_metric_card(slide, 6.25, 3.15, 2.0, "PATTERN", "Strategy", TEAL)
    add_metric_card(slide, 8.55, 3.15, 2.0, "VIEW", "Thymeleaf", BLUE)
    add_metric_card(slide, 10.85, 3.15, 1.82, "DATABASE", "MySQL", GOLD)
    add_footer(slide)

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(WHITE)
    add_header(slide, "Problem", "The challenge behind the model", 3)
    add_text(slide, "A parking slot is both a shared resource and a stateful object.", 0.65, 1.48, 11.7, 0.55, 23, NAVY, True)
    challenges = [
        ("01", "Limited capacity", "Prevent a slot being booked twice."),
        ("02", "Live occupancy", "Move consistently between available, booked and occupied."),
        ("03", "Payment truth", "Calculate duration and persist a payment outcome."),
        ("04", "Role visibility", "Expose different dashboards to authenticated roles."),
    ]
    for i, (num, title, desc) in enumerate(challenges):
        left = 0.65 + i * 3.05
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                      PInches(left), PInches(2.35),
                                      PInches(2.72), PInches(3.35))
        set_shape_fill(card, "F7FAFB")
        set_shape_line(card, "D7E4EA", 1)
        add_text(slide, num, left + 0.22, 2.62, 0.62, 0.46, 18, TEAL, True)
        add_text(slide, title, left + 0.22, 3.28, 2.22, 0.58, 17, NAVY, True)
        add_text(slide, desc, left + 0.22, 4.05, 2.22, 1.15, 13, "667B8B")
    add_footer(slide)

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(WHITE)
    add_header(slide, "Architecture", "Request to persistence", 4)
    stages = [
        ("1", "Browser", "HTTP"),
        ("2", "Controllers", "Session + routing"),
        ("3", "Services", "Business behavior"),
        ("4", "Repositories", "JPA queries"),
        ("5", "MySQL", "Persisted state"),
    ]
    for i, (num, title, sub) in enumerate(stages):
        left = 0.55 + i * 2.52
        circle = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                                        PInches(left), PInches(2.05),
                                        PInches(1.32), PInches(1.32))
        set_shape_fill(circle, TEAL if i % 2 == 0 else BLUE)
        set_shape_line(circle, WHITE, 2)
        add_text(slide, num, left, 2.45, 1.32, 0.35, 17, WHITE, True,
                 align=PP_ALIGN.CENTER)
        add_text(slide, title, left - 0.25, 3.55, 1.82, 0.42, 15, NAVY, True,
                 align=PP_ALIGN.CENTER)
        add_text(slide, sub, left - 0.25, 4.03, 1.82, 0.48, 10, "718495",
                 align=PP_ALIGN.CENTER)
        if i < 4:
            arrow = slide.shapes.add_shape(MSO_SHAPE.CHEVRON,
                                           PInches(left + 1.52), PInches(2.42),
                                           PInches(0.55), PInches(0.42))
            set_shape_fill(arrow, "CBDCE5")
            set_shape_line(arrow)
    add_text(slide, "Thymeleaf views render after controller decisions.",
             0.65, 5.25, 12.03, 0.55, 17, INK, True, align=PP_ALIGN.CENTER)
    add_footer(slide)

    add_diagram_slide(prs, 5, "Architecture", "Use case model", "use-case",
                      "Guests, users and administrators interact with distinct system goals.")
    add_diagram_slide(prs, 6, "Domain", "Object model", "class-diagram",
                      "Associations encode ownership and transaction relationships.")
    add_diagram_slide(prs, 7, "Persistence", "Entity relationship model", "entity-relationship",
                      "The database mirrors key JPA relationships with foreign keys.")
    add_diagram_slide(prs, 8, "Workflow", "Booking sequence", "sequence-booking",
                      "Search → select → validate → persist the reservation.")
    add_diagram_slide(prs, 9, "Workflow", "Payment sequence", "sequence-payment",
                      "Controller orchestration delegates method behavior to a strategy.")
    add_diagram_slide(prs, 10, "Workflow", "Entry and exit sequence", "sequence-entry-exit",
                      "Log timestamps and slot status keep the physical and digital states aligned.")
    add_diagram_slide(prs, 11, "Behavior", "Activity model", "activity",
                      "Decisions branch around authentication, availability and transaction completion.")
    add_diagram_slide(prs, 12, "Behavior", "State machine", "state-machine",
                      "Slot, booking and payment states change through explicit domain events.")
    add_diagram_slide(prs, 13, "System", "Component model", "component",
                      "Presentation, web, service, strategy and persistence packages collaborate.")
    add_diagram_slide(prs, 14, "System", "Deployment model", "deployment",
                      "Browser, application host, MySQL and SMTP form the runtime topology.")
    add_diagram_slide(prs, 15, "System", "Package model", "package",
                      "The source is organized by web, domain, persistence, service and strategy concerns.")

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(WHITE)
    add_header(slide, "OOP", "Four principles in executable code", 16)
    principles = [
        ("Encapsulation", "Private fields + getters/setters", TEAL),
        ("Abstraction", "Repository and strategy interfaces", BLUE),
        ("Polymorphism", "Card / Cash / Arrival strategies", GOLD),
        ("Composition", "Booking connects user, vehicle and slot", GREEN),
    ]
    for i, (title, evidence, color) in enumerate(principles):
        left = 0.65 + (i % 2) * 6.05
        top = 1.55 + (i // 2) * 2.48
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                      PInches(left), PInches(top),
                                      PInches(5.7), PInches(2.05))
        set_shape_fill(card, "F7FAFB")
        set_shape_line(card, "D7E4EA", 1)
        swatch = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                        PInches(left), PInches(top),
                                        PInches(0.12), PInches(2.05))
        set_shape_fill(swatch, color)
        set_shape_line(swatch)
        add_text(slide, title, left + 0.38, top + 0.32, 4.9, 0.45, 20, NAVY, True)
        add_text(slide, evidence, left + 0.38, top + 1.05, 4.9, 0.55, 14, "667B8B")
    add_footer(slide)

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(WHITE)
    add_header(slide, "Patterns", "Strategy pattern at runtime", 17)
    add_text(slide, "PaymentContext", 0.8, 2.0, 2.5, 0.55, 20, NAVY, True,
             align=PP_ALIGN.CENTER)
    context = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                     PInches(0.65), PInches(2.58),
                                     PInches(2.8), PInches(1.25))
    set_shape_fill(context, "EEF5F8")
    set_shape_line(context, "AFC5D1", 1)
    add_text(slide, "paymentStrategy\nprocessPayment()", 0.65, 2.78, 2.8, 0.72,
             15, INK, True, align=PP_ALIGN.CENTER)
    names = [("Card", "CardPaymentStrategy"), ("Cash", "CashPaymentStrategy"),
             ("Arrival", "ArrivalPaymentStrategy")]
    for i, (name, impl) in enumerate(names):
        top = 1.65 + i * 1.55
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                      PInches(7.15), PInches(top),
                                      PInches(4.55), PInches(1.05))
        set_shape_fill(card, "F7FAFB")
        set_shape_line(card, "D7E4EA", 1)
        add_text(slide, name, 7.42, top + 0.18, 1.1, 0.32, 15, TEAL, True)
        add_text(slide, impl, 8.45, top + 0.18, 3.0, 0.32, 13, NAVY, True)
        arrow = slide.shapes.add_shape(MSO_SHAPE.CHEVRON,
                                       PInches(5.05), PInches(top + 0.32),
                                       PInches(1.25), PInches(0.38))
        set_shape_fill(arrow, "CBDCE5")
        set_shape_line(arrow)
    add_text(slide, "The controller stays unchanged while the selected strategy changes.",
             3.75, 6.08, 5.85, 0.48, 14, "667B8B", True, align=PP_ALIGN.CENTER)
    add_footer(slide)

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(WHITE)
    add_header(slide, "SOLID", "An evidence-based assessment", 18)
    rows = [
        ("S", "Single Responsibility", "Layered web, service and persistence concerns", TEAL),
        ("O", "Open / Closed", "Add strategies without rewriting contexts", BLUE),
        ("L", "Liskov Substitution", "Concrete strategies satisfy interfaces", GOLD),
        ("I", "Interface Segregation", "Payment and authentication interfaces stay focused", GREEN),
        ("D", "Dependency Inversion", "Spring and contexts depend on abstractions", "9B78C6"),
    ]
    for i, (letter, title, evidence, color) in enumerate(rows):
        top = 1.48 + i * 0.97
        badge = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                                       PInches(0.75), PInches(top),
                                       PInches(0.62), PInches(0.62))
        set_shape_fill(badge, color)
        set_shape_line(badge)
        add_text(slide, letter, 0.75, top + 0.12, 0.62, 0.28, 15, WHITE, True,
                 align=PP_ALIGN.CENTER)
        add_text(slide, title, 1.68, top + 0.05, 3.15, 0.38, 16, NAVY, True)
        add_text(slide, evidence, 4.95, top + 0.05, 7.55, 0.38, 14, "667B8B")
    add_footer(slide)

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(WHITE)
    add_header(slide, "Quality", "What is verified and what is still open", 19)
    add_metric_card(slide, 0.65, 1.52, 2.72, "LAYERING", "Separated", TEAL)
    add_metric_card(slide, 3.62, 1.52, 2.72, "MODELING", "11 UML views", BLUE)
    add_metric_card(slide, 6.59, 1.52, 2.72, "STRATEGIES", "Payment + role", GOLD)
    add_metric_card(slide, 9.56, 1.52, 3.11, "TEST BASIS", "Maven + Surefire", GREEN)
    add_rich_text(slide, [
        "Add concurrent-slot locking",
        "Expand integration test coverage",
        "Replace dashboard placeholder uptime",
        "Harden authentication and password storage",
        "Connect real telemetry and monitoring",
    ], 0.65, 3.22, 7.2, 2.55, 16)
    add_footer(slide)

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(NAVY)
    add_text(slide, "CONCLUSION", 0.85, 0.72, 3.2, 0.38, 13, CYAN, True)
    add_text(slide, "A coherent model for\nreal-world parking operations",
             0.85, 1.55, 8.2, 1.42, 31, WHITE, True)
    add_text(slide,
             "Objects, states and strategies turn a shared resource into a maintainable, understandable system.",
             0.85, 3.35, 7.8, 0.82, 18, "C9D7E3")
    add_text(slide, "Thank you", 9.18, 5.35, 3.0, 0.58, 28, GOLD, True,
             align=PP_ALIGN.CENTER)
    add_text(slide, "Questions & discussion", 8.85, 6.12, 3.65, 0.42, 14,
             WHITE, True, align=PP_ALIGN.CENTER)

    for i, slide in enumerate(prs.slides):
        if i == 0:
            continue
        if i < len(prs.slides) - 1:
            next_button = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                PInches(11.72), PInches(6.82), PInches(0.96), PInches(0.38),
            )
            set_shape_fill(next_button, NAVY)
            set_shape_line(next_button)
            next_button.click_action.target_slide = prs.slides[i + 1]
            add_text(slide, "NEXT >", 11.72, 6.86, 0.96, 0.22, 9, WHITE, True,
                     align=PP_ALIGN.CENTER, margin=0)

        if i > 1:
            back_button = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                PInches(10.58), PInches(6.82), PInches(0.96), PInches(0.38),
            )
            set_shape_fill(back_button, "DCE7EC")
            set_shape_line(back_button)
            back_button.click_action.target_slide = prs.slides[i - 1]
            add_text(slide, "< BACK", 10.58, 6.86, 0.96, 0.22, 9, NAVY, True,
                     align=PP_ALIGN.CENTER, margin=0)

        if i > 1:
            home_button = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                PInches(9.44), PInches(6.82), PInches(0.96), PInches(0.38),
            )
            set_shape_fill(home_button, "EEF5F8")
            set_shape_line(home_button)
            home_button.click_action.target_slide = prs.slides[0]
            add_text(slide, "HOME", 9.44, 6.86, 0.96, 0.22, 9, BLUE, True,
                     align=PP_ALIGN.CENTER, margin=0)

    prs.save(PRESENTATION / "ParQHub_OOPS_Presentation.pptx")


if __name__ == "__main__":
    export_diagrams()
    generate_report()
    generate_presentation()

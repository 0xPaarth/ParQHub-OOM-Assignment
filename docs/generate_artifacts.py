from pathlib import Path
import subprocess

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from pptx import Presentation
from pptx.util import Inches as PInches, Pt as PPt


ROOT = Path(__file__).resolve().parent
UML = ROOT / "uml"
REPORT = ROOT / "report"
PRESENTATION = ROOT / "presentation"
REPORT.mkdir(exist_ok=True)
PRESENTATION.mkdir(exist_ok=True)


DOT = {
    "use-case": r'''digraph G { rankdir=LR; node [shape=ellipse, style=filled, fillcolor="#E8F1FF", fontname="Arial"]; Guest[label="Guest"]; User[label="User"]; Admin[label="Admin"]; subgraph cluster_app { label="ParQHub web application"; color="#9BB7D4"; Browse[label="Browse home/login"]; Register[label="Register user and vehicles"]; Login[label="Login"]; Search[label="Search parking"]; Book[label="Book slot"]; Vehicle[label="Entry/exit logging"]; Pay[label="Process payment"]; Notify[label="Notifications"]; Settings[label="Settings"]; Dashboard[label="Admin dashboard"]; Manage[label="Manage parking/tables"]; Reports[label="Reports"]; } Guest->Browse; Guest->Register; Guest->Login; User->Search; User->Book; User->Vehicle; User->Pay; User->Notify; User->Settings; Admin->Dashboard; Admin->Manage; Admin->Reports; }''',
    "class-diagram": r'''digraph G { rankdir=LR; node [shape=record, fontname="Arial", fontsize=10]; User[label="{User|id; name; email; password; contactNo}"]; Admin[label="{Admin|id; name; email; password; role}"]; Vehicle[label="{Vehicle|id; vehicleNo; vehicleType; brand; model; color}"]; Lot[label="{ParkingLot|id; city; location; totalSlots; priceHr}"]; Slot[label="{ParkingSlot|id; status}"]; Booking[label="{Booking|id; startTime; endTime; paymentStatus}"]; Payment[label="{Payment|id; amount; method; status; timestamp}"]; Log[label="{VehicleLog|id; entryTime; exitTime}"]; Notification[label="{Notification|id; type; description; timestamp}"]; Report[label="{Report|id; type; generatedDate}"]; User->Vehicle[label="1..*"]; User->Booking[label="1..*"]; User->Notification[label="1..*"]; Lot->Slot[label="1..*"]; Booking->User[label="*..1"]; Booking->Vehicle[label="*..1"]; Booking->Slot[label="*..1"]; Payment->Booking[label="*..1"]; Log->Vehicle[label="*..1"]; Log->Lot[label="*..1"]; Notification->User[label="*..0..1"]; Notification->Admin[label="*..0..1"]; Report->Admin[label="*..1"]; }''',
    "activity": r'''digraph G { rankdir=TB; node [shape=box, style="rounded,filled", fillcolor="#F4F8FC", fontname="Arial"]; start[shape=circle,label="",width=.2]; login[shape=diamond,label="Logged in?"]; search[label="Search city/location"]; slots[label="Load available slots"]; select[label="Select vehicle and slot"]; available[shape=diamond,label="Slot available?"]; book[label="Create pending booking"]; entry[label="Record entry; slot OCCUPIED"]; exitStep[label="Record exit"]; duration[label="Calculate duration"]; pay[label="Create/save payment"]; done[label="Complete booking; slot AVAILABLE"]; end[shape=doublecircle,label="",width=.2]; start->login; login->search[label="yes"]; login->login[label="no: show login"]; search->slots->select->available; available->search[label="no"]; available->book[label="yes"]; book->entry->exitStep->duration->pay->done->end; }''',
    "state-machine": r'''digraph G { rankdir=LR; node [shape=ellipse, style=filled, fillcolor="#FFF8E6", fontname="Arial"]; start[shape=point]; available[label="AVAILABLE"]; booked[label="BOOKED"]; occupied[label="OCCUPIED"]; start->available; available->booked[label="booking created"]; booked->occupied[label="vehicle entry"]; occupied->available[label="vehicle exit"]; booked->available[label="completion/cancel flow"]; pending[label="Booking Pending"]; completed[label="Booking Completed"]; pending->completed[label="payment completed"]; pp[label="Payment Pending"]; pc[label="Payment Completed"]; pp->pc[label="Card/Cash"]; }''',
    "component": r'''digraph G { rankdir=LR; node [shape=box, style="rounded,filled", fillcolor="#EAF7EE", fontname="Arial"]; browser[label="Browser"]; mvc[label="Spring MVC controllers"]; views[label="Thymeleaf templates"]; services[label="Service layer"]; strategies[label="Payment/admin-auth strategies"]; repos[label="Spring Data JPA repositories"]; db[shape=cylinder,label="MySQL parqhub_db"]; mail[label="JavaMailSender"]; browser->mvc; mvc->views; mvc->services; mvc->repos; services->strategies; services->repos; mvc->mail; repos->db; }''',
    "deployment": r'''digraph G { rankdir=LR; node [shape=box3d, style=filled, fillcolor="#EEF2FF", fontname="Arial"]; client[label="User/Admin browser"]; app[label="Application host\nJDK 21 + Spring Boot 3.3.4\nHTTP 8080"]; db[shape=cylinder,label="MySQL parqhub_db"]; smtp[label="SMTP mail service"]; client->app[label="HTTP"]; app->db[label="JDBC"]; app->smtp[label="SMTP"]; }''',
    "package": r'''digraph G { rankdir=TB; node [shape=folder, style=filled, fillcolor="#F4F4F4", fontname="Arial"]; root[label="com.prmplatform.parqhub"]; controller[label="controller"]; admin[label="controller.admin"]; user[label="controller.user"]; model[label="model"]; repo[label="repository"]; service[label="service"]; strategy[label="strategy"]; payment[label="strategy.payment"]; auth[label="strategy.adminauth"]; root->controller; root->model; root->repo; root->service; root->strategy; controller->admin; controller->user; strategy->payment; strategy->auth; controller->service; controller->repo; service->repo; service->strategy; repo->model; }''',
    "entity-relationship": r'''digraph G { rankdir=LR; node [shape=record, fontname="Arial", fontsize=9]; Users[label="{USERS|user_id PK|name|email UK|password|contact_no}"]; Admins[label="{ADMINS|admin_id PK|name|email UK|password|role}"]; Vehicle[label="{VEHICLE|vehicle_id PK|user_id FK|vehicle_no UK|vehicle_type}"]; Lot[label="{PARKINGLOT|lot_id PK|city|location|total_slots|price_hr}"]; Slot[label="{PARKINGSLOT|slot_id PK|lot_id FK|status}"]; Booking[label="{BOOKING|booking_id PK|user_id FK|vehicle_id FK|slot_id FK|start_time|end_time|payment_status}"]; Payment[label="{PAYMENT|payment_id PK|booking_id FK|amount|method|status|timestamp}"]; Log[label="{VEHICLELOG|log_id PK|vehicle_id FK|lot_id FK|entry_time|exit_time}"]; Report[label="{REPORT|report_id PK|admin_id FK|type|generated_date}"]; Notification[label="{NOTIFICATION|notification_id PK|user_id FK|admin_id FK|type|description|timestamp}"]; Users->Vehicle; Users->Booking; Users->Notification; Admins->Report; Admins->Notification; Lot->Slot; Vehicle->Booking; Slot->Booking; Booking->Payment; Vehicle->Log; Lot->Log; }''',
    "sequence-booking": r'''digraph G { rankdir=LR; node [shape=box, style=filled, fillcolor="#F4F8FC", fontname="Arial"]; U[label="User"]; UI[label="Thymeleaf UI"]; C[label="UserController"]; V[label="VehicleRepository"]; S[label="ParkingSlotRepository"]; B[label="BookingRepository"]; D[label="MySQL"]; U->UI->C->V; C->S; C->B; V->D; S->D; B->D; }''',
    "sequence-payment": r'''digraph G { rankdir=LR; node [shape=box, style=filled, fillcolor="#F4F8FC", fontname="Arial"]; U[label="User"]; C[label="PaymentController"]; P[label="PaymentProcessingService"]; X[label="PaymentContext"]; T[label="Card/Cash/Arrival strategy"]; R[label="Repositories"]; M[label="JavaMailSender"]; U->C->P->X->T; C->R; C->M; }''',
    "sequence-entry-exit": r'''digraph G { rankdir=LR; node [shape=box, style=filled, fillcolor="#F4F8FC", fontname="Arial"]; U[label="User"]; C[label="VehicleLogController"]; B[label="BookingRepository"]; L[label="VehicleLogRepository"]; S[label="ParkingSlotRepository"]; U->C->B; C->L; C->S; }''',
}


def export_diagrams():
    for name, source in DOT.items():
        dot_file = UML / f".{name}.dot"
        dot_file.write_text(source)
        for fmt in ("svg", "png"):
            subprocess.run(["dot", f"-T{fmt}", str(dot_file), "-o", str(UML / f"{name}.{fmt}")], check=True)
        dot_file.unlink()


def add_title(doc, text, size=24):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(size)


def add_heading(doc, text, level=1):
    doc.add_heading(text, level=level)


def add_bullets(doc, items):
    for item in items:
        doc.add_paragraph(item, style="List Bullet")


def generate_report():
    doc = Document()
    add_title(doc, "ParQHub")
    add_title(doc, "Object-Oriented Modelling and OOPS Report", 18)
    doc.add_paragraph("\nTeam Members", style="Title")
    for member in ("Paarth Arora", "Mannat Jain", "Kajal Jayaswal"):
        doc.add_paragraph(member)
    doc.add_page_break()

    sections = [
        ("1. Introduction and Problem Statement", [
            "ParQHub is a Spring Boot web application for vehicle parking reservation and management. It supports user registration, vehicle management, parking search, slot booking, entry and exit logging, payment processing, notifications, administration, and reports.",
            "The problem addressed is the coordination of limited parking capacity, reservations, occupancy transitions, payment records, and administrative monitoring through a single web application."
        ]),
        ("2. Requirements", [
            "Functional requirements: register users and vehicles; authenticate users and administrators; search lots by city and location; book available slots; record vehicle entry and exit; calculate parking duration; process Card, Cash, and Arrival payments; view history and notifications; manage parking data; generate reports.",
            "Non-functional requirements evidenced by the implementation: layered separation of controllers, services, repositories, and views; persistence through JPA; role-specific administration; transactional parking-slot updates; browser-based access through Spring MVC and Thymeleaf."
        ]),
        ("3. System Architecture", [
            "The application follows a Spring MVC structure. Controllers handle HTTP requests and return Thymeleaf views or API responses. Services encapsulate business operations. Spring Data JPA repositories abstract persistence. Model classes are JPA entities. Payment and administrator authentication use explicit Strategy and Context classes.",
            "The configured runtime is Java 21 with Spring Boot 3.3.4, MySQL through the MySQL connector, Thymeleaf, Spring Mail, and Maven."
        ]),
        ("4. Object Model and Relationships", [
            "User owns vehicles, bookings, and notifications. ParkingLot contains ParkingSlot instances. A Booking references a User, Vehicle, and ParkingSlot. Payment references a Booking. VehicleLog references a Vehicle and ParkingLot. Report references an Admin. Notification may reference a User and/or Admin.",
            "DTO classes ParkingSlotDTO and VehicleDTO provide non-persistent transfer shapes."
        ]),
        ("5. OOP Concepts Demonstrated", [
            "Encapsulation is represented by private entity fields with public getters and setters.",
            "Abstraction is represented by repository interfaces, service APIs, and PaymentStrategy/AdminAuthenticationStrategy interfaces.",
            "Polymorphism is used when PaymentContext delegates to CardPaymentStrategy, CashPaymentStrategy, or ArrivalPaymentStrategy, and when the authentication context delegates to role-specific strategies.",
            "Composition and association are represented by JPA relationships such as User-to-Vehicle, ParkingLot-to-ParkingSlot, and Booking-to-Payment.",
            "Inheritance is not a dominant feature of the domain model; the implementation favors interfaces and composition."
        ]),
        ("6. Design Patterns and SOLID Discussion", [
            "MVC, Repository, Service Layer, Dependency Injection, DTO, and Strategy patterns are directly evidenced by the source structure and annotations.",
            "Single Responsibility is supported by separating request handling, business services, persistence, and strategy selection. Open/Closed is supported by adding payment strategies behind PaymentStrategy. Dependency Inversion is supported through interfaces and Spring-managed dependencies. Interface Segregation is represented by focused strategy interfaces. Liskov Substitution applies to concrete strategy implementations used through their interfaces.",
            "These are architectural observations based on the current code, not claims that every class is perfectly compliant."
        ]),
        ("7. Workflow and State Analysis", [
            "Booking moves a slot from AVAILABLE to BOOKED. Vehicle entry moves it to OCCUPIED. Vehicle exit and the payment completion flow return it to AVAILABLE. Booking and payment statuses use Pending, Completed, and Failed values. A VehicleLog is active when exitTime is null and completed after exitTime is set."
        ]),
        ("8. Conclusion", [
            "ParQHub provides a coherent object-oriented parking-management implementation with a layered Spring architecture, JPA domain model, explicit strategies, and role-aware workflows. The accompanying diagrams document the actual source code and database schema for academic analysis."
        ]),
    ]
    for title, paragraphs in sections:
        add_heading(doc, title)
        for paragraph in paragraphs:
            doc.add_paragraph(paragraph)

    add_heading(doc, "9. UML Diagram Index")
    add_bullets(doc, [
        "Use Case Diagram: actors and supported user/admin goals.",
        "Class Diagram: entity classes, DTOs, fields, and associations.",
        "Sequence Diagrams: booking, payment, and vehicle entry/exit workflows.",
        "Activity Diagram: user parking-to-payment flow.",
        "State Machine Diagram: parking-slot, booking, payment, and vehicle-log states.",
        "Component, Deployment, Package, and Entity Relationship diagrams: implementation structure and runtime/data deployment."
    ])
    for name in ("use-case", "class-diagram", "sequence-booking", "sequence-payment", "sequence-entry-exit", "activity", "state-machine", "component", "deployment", "package", "entity-relationship"):
        image = UML / f"{name}.png"
        if image.exists():
            doc.add_heading(name.replace("-", " ").title(), level=2)
            doc.add_picture(str(image), width=Inches(6.2))
            doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.save(REPORT / "ParQHub_OOM_Report.docx")


def generate_presentation():
    prs = Presentation()
    prs.slide_width = PInches(13.333)
    prs.slide_height = PInches(7.5)

    def slide(title, body=None, image=None):
        layout = prs.slide_layouts[6]
        s = prs.slides.add_slide(layout)
        tx = s.shapes.add_textbox(PInches(0.55), PInches(0.3), PInches(12.2), PInches(0.7))
        p = tx.text_frame.paragraphs[0]
        p.text = title
        p.font.size = PPt(28)
        p.font.bold = True
        if body:
            box = s.shapes.add_textbox(PInches(0.75), PInches(1.2), PInches(11.8), PInches(5.4))
            tf = box.text_frame
            tf.word_wrap = True
            for i, line in enumerate(body):
                para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                para.text = line
                para.font.size = PPt(20)
                para.space_after = PPt(12)
        if image and image.exists():
            s.shapes.add_picture(str(image), PInches(0.85), PInches(1.15), width=PInches(11.5), height=PInches(5.8))

    slide("ParQHub", ["Object-Oriented Modelling and OOPS Assignment", "Paarth Arora | Mannat Jain | Kajal Jayaswal"])
    slide("Project Overview", ["Spring Boot parking reservation and management platform.", "Supports users, administrators, parking lots, slots, bookings, payments, vehicle logs, notifications, and reports."])
    slide("Problem Statement", ["Parking capacity and reservations must be coordinated across users, vehicles, slots, payments, and administrative operations.", "ParQHub centralizes these workflows in one web application."])
    slide("Technology Stack", ["Java 21; Spring Boot 3.3.4; Spring MVC; Thymeleaf; Spring Data JPA/Hibernate; MySQL; Spring Mail; Maven; H2 runtime for testing."])
    slide("Architecture", ["Browser -> Spring MVC controllers -> services/repositories -> MySQL.", "Thymeleaf renders views; strategy contexts provide replaceable payment and admin-authentication behavior."])
    slide("Use Case Diagram", image=UML / "use-case.png")
    slide("Class Diagram", image=UML / "class-diagram.png")
    slide("Entity Relationship Diagram", image=UML / "entity-relationship.png")
    slide("Booking Workflow", image=UML / "sequence-booking.png")
    slide("Payment Workflow", image=UML / "sequence-payment.png")
    slide("Vehicle Entry and Exit", image=UML / "sequence-entry-exit.png")
    slide("Activity Model", image=UML / "activity.png")
    slide("State Machine", image=UML / "state-machine.png")
    slide("Component Model", image=UML / "component.png")
    slide("Deployment Model", image=UML / "deployment.png")
    slide("Package Model", image=UML / "package.png")
    slide("OOP Concepts", ["Encapsulation through private entity fields and accessors.", "Abstraction through repository and strategy interfaces.", "Polymorphism through payment and admin-authentication strategies.", "Composition through JPA associations."])
    slide("Design Patterns", ["MVC; Repository; Service Layer; Dependency Injection; DTO; Strategy.", "Strategy is explicitly implemented with contexts and interchangeable concrete strategies."])
    slide("SOLID Discussion", ["Focused layers support Single Responsibility.", "Interfaces and strategies support Open/Closed and Dependency Inversion.", "Concrete strategies are substitutable through their interfaces.", "The assessment documents the current implementation rather than asserting perfection."])
    slide("Conclusion", ["ParQHub demonstrates a practical object-oriented web architecture for parking management.", "The UML artifacts reflect the current Java source and SQL schema."])
    prs.save(PRESENTATION / "ParQHub_OOPS_Presentation.pptx")


if __name__ == "__main__":
    export_diagrams()
    generate_report()
    generate_presentation()

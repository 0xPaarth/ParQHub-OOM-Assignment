# ParQHub - Smart Parking Management System

ParQHub is a Java 21 and Spring Boot web application for vehicle parking reservation and management. The current implementation supports user registration, vehicle management, parking search, slot booking, vehicle entry/exit logging, payment processing, notifications, administration, and operational reports.

## Team Members

- Paarth Arora
- Mannat Jain
- Kajal Jayaswal

## Features

### User workflows

- Register a user with one or more vehicles.
- Log in using the user session.
- Search parking lots by city and location.
- View available parking slots and create bookings.
- Record vehicle entry and exit.
- Calculate parking duration and process Card, Cash, or Arrival payments.
- View bookings, payments, downloadable text receipts, notifications, and settings.

### Administrator workflows

- Authenticate administrators by role.
- View role-filtered dashboard information.
- Inspect parking lots, slots, and occupancy statistics.
- Manage application tables through the administrative interface.
- Generate financial, occupancy, and performance reports.
- Send parking-slot/security notifications.

## Technology Stack

- Java 21
- Spring Boot 3.3.4
- Spring MVC and Thymeleaf
- Spring Data JPA / Hibernate
- MySQL connector 8.1.0
- Spring Mail
- Maven
- H2 runtime dependency for testing

The configured datasource is `jdbc:mysql://localhost:3306/parqhub_db`. Application settings are in [`src/main/resources/application.properties`](src/main/resources/application.properties). The repository does not currently include the Spring Security starter; authentication is implemented with repository lookups and HTTP sessions.

## Architecture

```text
Browser
  -> Spring MVC controllers
  -> Service layer and Strategy contexts
  -> Spring Data JPA repositories
  -> MySQL parqhub_db
```

- `controller`: public, user, and administrator request handlers.
- `model`: JPA entities and the `ParkingSlotDTO` / `VehicleDTO` transfer classes.
- `repository`: Spring Data JPA repository interfaces.
- `service`: parking-slot, parking-viewer, payment, and administrator-authentication services.
- `strategy/payment`: Card, Cash, and Arrival payment strategies.
- `strategy/adminauth`: role-specific administrator-authentication strategies.
- `src/main/resources/templates`: Thymeleaf views.

## OOP and Design Patterns

The implementation demonstrates:

- **Encapsulation:** private model fields with public accessors.
- **Abstraction:** repository and strategy interfaces.
- **Polymorphism:** interchangeable payment and administrator-authentication strategies.
- **Composition and association:** JPA relationships between users, vehicles, bookings, slots, payments, logs, reports, and notifications.
- **MVC:** controllers, model entities, and Thymeleaf views.
- **Repository pattern:** `JpaRepository` interfaces abstract persistence.
- **Service Layer pattern:** business operations are separated from request handling.
- **Dependency Injection:** Spring-managed dependencies are injected into controllers and services.
- **Strategy pattern:** `PaymentContext` and `AdminAuthenticationContext` delegate to concrete strategies.
- **DTO pattern:** `ParkingSlotDTO` and `VehicleDTO` expose focused transfer data.

The SOLID discussion in the report is evidence-based: the code supports separation of responsibilities, interfaces for replaceable strategies, and dependency inversion through Spring-managed abstractions. It does not claim that every class is perfectly compliant.

## UML Documentation

Mermaid source files and exported SVG/PNG diagrams are in [`docs/uml`](docs/uml):

| Diagram | Source | Export |
|---|---|---|
| Use Case | [`use-case.mmd`](docs/uml/use-case.mmd) | [`use-case.svg`](docs/uml/use-case.svg) |
| Class | [`class-diagram.mmd`](docs/uml/class-diagram.mmd) | [`class-diagram.svg`](docs/uml/class-diagram.svg) |
| Booking sequence | [`sequence-booking.mmd`](docs/uml/sequence-booking.mmd) | [`sequence-booking.svg`](docs/uml/sequence-booking.svg) |
| Payment sequence | [`sequence-payment.mmd`](docs/uml/sequence-payment.mmd) | [`sequence-payment.svg`](docs/uml/sequence-payment.svg) |
| Entry/exit sequence | [`sequence-entry-exit.mmd`](docs/uml/sequence-entry-exit.mmd) | [`sequence-entry-exit.svg`](docs/uml/sequence-entry-exit.svg) |
| Activity | [`activity.mmd`](docs/uml/activity.mmd) | [`activity.svg`](docs/uml/activity.svg) |
| State machine | [`state-machine.mmd`](docs/uml/state-machine.mmd) | [`state-machine.svg`](docs/uml/state-machine.svg) |
| Component | [`component.mmd`](docs/uml/component.mmd) | [`component.svg`](docs/uml/component.svg) |
| Deployment | [`deployment.mmd`](docs/uml/deployment.mmd) | [`deployment.svg`](docs/uml/deployment.svg) |
| Package | [`package.mmd`](docs/uml/package.mmd) | [`package.svg`](docs/uml/package.svg) |
| Entity relationship | [`entity-relationship.mmd`](docs/uml/entity-relationship.mmd) | [`entity-relationship.svg`](docs/uml/entity-relationship.svg) |

The diagrams are based on the Java entities, controller mappings, service/strategy classes, repository interfaces, and SQL foreign keys in this repository.

## Assignment Deliverables

- [Object-Oriented Modelling report - DOCX](docs/report/ParQHub_OOM_Report.docx)
- [Object-Oriented Modelling report - PDF](docs/report/ParQHub_OOM_Report.pdf)
- [OOPS presentation - PPTX](docs/presentation/ParQHub_OOPS_Presentation.pptx)

## Run the Application

Prerequisites: JDK 21, Maven, and a MySQL-compatible database.

```bash
./mvnw test
./mvnw spring-boot:run
```

Create/configure the `parqhub_db` database before starting the application. The SQL reference files are [`db.sql`](db.sql) and [`crdb.sql`](crdb.sql).

## Validation

The generated Office documents can be opened with Microsoft Office or LibreOffice. The PDF is exported from the DOCX using LibreOffice. The application source code is unchanged; documentation and assignment deliverables are under `docs/`, with this README updated to describe the actual implementation.

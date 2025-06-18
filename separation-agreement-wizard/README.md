# Separation Agreement Wizard

A comprehensive web application designed to streamline the creation of legal separation agreements. This wizard guides users through a step-by-step process to collect necessary information and generate customized legal documents.

## Project Overview

The Separation Agreement Wizard simplifies the complex process of creating legal separation agreements by:

1. Collecting information about both parties involved
2. Allowing customization of agreement clauses based on specific circumstances
3. Generating a professionally formatted legal document in DOCX format

## Technology Stack

### Frontend
- **Framework**: Angular 19
- **UI Components**: PrimeNG 19.1.0
- **Styling**: PrimeFlex 3.3.1, SCSS
- **State Management**: RxJS with BehaviorSubject
- **Form Handling**: Angular Reactive Forms

### Backend
- **.NET Core**: .NET 9.0 Web API
- **Document Generation**: DocumentFormat.OpenXml 3.3.0
- **API Documentation**: Swagger/OpenAPI

## Application Architecture

The application follows a client-server architecture:

### Frontend Components
- **Party Information Form**: Collects personal details of both parties and any children
- **Clause Selection**: Allows users to select and customize legal clauses by category
- **Document Review**: Provides a preview and download option for the generated document

### Backend Services
- **Document Controller**: Handles API requests for document generation
- **Document Generation Service**: Creates formatted DOCX files based on user inputs

## Features

- Multi-step wizard interface with form validation
- Persistent session storage to prevent data loss during navigation
- Categorized legal clauses with customization options
- Dynamic document generation with proper legal formatting
- Responsive design for desktop and mobile use

## Data Flow

1. User enters party information (names, roles, dates, children)
2. User selects relevant legal clauses from predefined categories
3. Backend processes the request and generates a properly formatted legal document
4. User can download the document as a DOCX file

## Getting Started

### Development Environment Setup

```bash
# Clone the repository
git clone [repository-url]

# Navigate to the Angular project directory
cd separation-agreement-wizard

# Install dependencies
npm install

# Start the Angular development server
ng serve

# In a separate terminal, navigate to the API directory
cd ../SeparationAgreementApi/DocumentGenerationApi

# Run the .NET API
dotnet run
```

The application will be available at `http://localhost:4200/` and will automatically connect to the API running at `https://localhost:7227/`.

## Building for Production

```bash
# Build the Angular application
ng build --configuration production

# Publish the .NET API
dotnet publish -c Release
```

## Testing

```bash
# Run Angular unit tests
ng test

# Run .NET API tests
dotnet test
```

## Additional Resources

For more information on using the Angular CLI, including detailed command references, visit the [Angular CLI Overview and Command Reference](https://angular.dev/tools/cli) page.

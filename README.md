# Duty Rotation Scheduler

This is a web application for managing duty schedules, built with React and TypeScript.

## Features

*   **Customizable Schedule:**
    *   Editable application title.
    *   Selectable schedule start month.
    *   Selectable start day of the week.
    *   **Strict Calendar Year Display:** The schedule view only shows weeks starting within the selected calendar year (e.g., selecting 2025 and July start month shows July 2025 - Dec 2025).
    *   Selectable year (configurable range, e.g., 2019-2031).
*   **Team Management:**
    *   Add and remove team members (up to 20).
    *   Edit team member names.
    *   Manage leave periods for each team member.
*   **Holiday Integration:**
    *   Automatically fetches US public holidays from the Nager.Date API for the selected year and the following year (to handle year-end period).
    *   Highlights weeks containing official public holidays.
    *   **Precise Holiday Period:** Automatically marks weeks falling within the defined holiday period (from the week containing the Monday *before* Christmas through the week containing the Friday *after* New Year's Day) as "Holiday Period" and leaves them unassigned.
*   **Fair Assignment:**
    *   Assigns duties based on availability (considering leave and the holiday period).
    *   Attempts to balance total duties and holiday duties among team members.
    *   Allows manual override of assignments (outside the holiday period).
*   **Pop-up Notes:**
    *   Add, view, and edit text notes for each week.
*   **CSV Import/Export:**
    *   Export the current schedule (Week, Assigned To, Holidays, Notes) to a CSV file.
    *   Import a previously exported schedule from a CSV file (Note: Team members and leave are not imported/exported via CSV, only the schedule itself. Imported assignments during the holiday period will be overridden).
*   **Analytics:**
    *   Displays a summary table showing total duties and holiday duties assigned per person for the generated schedule.

## Setup and Running Locally

1.  **Prerequisites:** Node.js and pnpm (or npm/yarn).
2.  **Install Dependencies:**
    ```bash
    pnpm install
    ```
3.  **Run Development Server:**
    ```bash
    pnpm run dev
    ```
    The application will be available at `http://localhost:5173` (or another port if 5173 is busy).

4.  **Build for Production:**
    ```bash
    pnpm run build
    ```
    The static files will be generated in the `dist` directory.

## Project Structure

*   `public/`: Static assets.
*   `src/`: Source code.
    *   `components/`: React components (includes `duty-rotation-scheduler.tsx`).
    *   `App.tsx`: Main application component.
    *   `main.tsx`: Application entry point.
    *   `index.css`: Global styles (Tailwind CSS).
*   `dist/`: Production build output.
*   `package.json`: Project metadata and dependencies.
*   `pnpm-lock.yaml`: Lockfile for dependencies.
*   `vite.config.ts`: Vite build configuration.
*   `tsconfig.json`: TypeScript configuration.

## Deployment

This application is built as a static site. The contents of the `dist` directory can be deployed to any static hosting provider (like GitHub Pages, Vercel, Netlify, etc.).




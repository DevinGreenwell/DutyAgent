import React, { useState, useEffect, useCallback } from "react";
import {
  Trash2,
  Plus,
  Upload,
  Edit,
  Save,
  X,
  StickyNote,
  FileText
} from "lucide-react";

// --- Type Definitions ---
interface Leave {
  start: string; // YYYY-MM-DD
  end: string;   // YYYY-MM-DD
}

interface Person {
  id: number;
  name: string;
  leave: Leave[];
}

interface Holiday {
  date: string; // YYYY-MM-DD
  name: string;
}

interface Week {
  startDate: Date;
  endDate: Date;
  assignedTo: number | string | null; // Person ID, "Holiday Period", "No one available", or null
  hasHoliday: boolean;
  isHolidayPeriod: boolean;
  notes: string;
}

interface HolidayCounts {
  [personId: number]: {
    total: number;
    holiday: number;
    lastAssigned?: number; // Optional, used during assignment logic
  };
}

// --- Helper Constants ---
const dayNames = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const MAX_PEOPLE = 20;
const SCHEDULE_YEAR_RANGE = 13; // e.g., currentYear - 5 to currentYear + 7

// CSV Header Constants (Ensure these match export/import logic)
const CSV_HEADERS = [
    "Week Start Date", // YYYY-MM-DD
    "Week End Date",   // YYYY-MM-DD
    "Assigned To",     // Person Name or Status String
    "Holidays",        // Multi-line string: "YYYY-MM-DD: Name\n..."
    "Notes"            // Multi-line string
];

// Helper function to get the Monday before a given date
const getMondayBefore = (date: Date): Date => {
    const dayOfWeek = date.getUTCDay(); // 0 = Sunday, 1 = Monday, ...
    const diff = dayOfWeek === 0 ? -6 : 1 - dayOfWeek; // Calculate days to subtract to get to Monday
    const monday = new Date(date);
    monday.setUTCDate(date.getUTCDate() + diff);
    return monday;
};

// Helper function to get the Friday after a given date
const getFridayAfter = (date: Date): Date => {
    const dayOfWeek = date.getUTCDay(); // 0 = Sunday, ..., 5 = Friday, 6 = Saturday
    const diff = dayOfWeek === 6 ? 6 : 5 - dayOfWeek; // Calculate days to add to get to Friday
    const friday = new Date(date);
    friday.setUTCDate(date.getUTCDate() + diff);
    return friday;
};

const DutyRotationScheduler: React.FC = () => {
  // --- State Variables ---
  const [appTitle, setAppTitle] = useState<string>(() => localStorage.getItem('dutySchedulerAppTitle') || "Duty Rotation Scheduler");
  const [isEditingTitle, setIsEditingTitle] = useState<boolean>(false);
  const [tempTitle, setTempTitle] = useState<string>(appTitle);

  const currentYear = new Date().getFullYear();
  const [year, setYear] = useState<number>(currentYear);
  const [startDayOfWeek, setStartDayOfWeek] = useState<number>(() => parseInt(localStorage.getItem('dutySchedulerStartDay') || '2', 10)); // Default Tuesday (2)
  const [startMonth, setStartMonth] = useState<number>(() => parseInt(localStorage.getItem('dutySchedulerStartMonth') || '0', 10)); // Default January (0)

  const [people, setPeople] = useState<Person[]>(() => {
    const savedPeople = localStorage.getItem('dutySchedulerPeople');
    try {
        return savedPeople ? JSON.parse(savedPeople) : [
          { id: 1, name: "Person 1", leave: [] },
          { id: 2, name: "Person 2", leave: [] },
          { id: 3, name: "Person 3", leave: [] }
        ];
    } catch (e) {
        console.error("Failed to parse people from localStorage", e);
        return [
          { id: 1, name: "Person 1", leave: [] },
          { id: 2, name: "Person 2", leave: [] },
          { id: 3, name: "Person 3", leave: [] }
        ];
    }
  });
  const [newPersonName, setNewPersonName] = useState<string>("");
  const [schedule, setSchedule] = useState<Week[]>([]);
  const [holidays, setHolidays] = useState<Holiday[]>([]);
  const [isLoadingHolidays, setIsLoadingHolidays] = useState<boolean>(false);
  const [holidayError, setHolidayError] = useState<string | null>(null);
  const [holidayCounts, setHolidayCounts] = useState<HolidayCounts>({});

  // Leave Form State
  const [leaveDate, setLeaveDate] = useState<string>("");
  const [leaveEndDate, setLeaveEndDate] = useState<string>("");
  const [selectedPersonIdForLeave, setSelectedPersonIdForLeave] = useState<number | null>(null);
  const [showLeaveForm, setShowLeaveForm] = useState<boolean>(false);

  // Notes Modal State
  const [showNotesModal, setShowNotesModal] = useState<boolean>(false);
  const [currentWeekIndexForNotes, setCurrentWeekIndexForNotes] = useState<number | null>(null);
  const [currentNotes, setCurrentNotes] = useState<string>("");

  // --- Effects --- 

  // Save settings and people to localStorage
  useEffect(() => {
    localStorage.setItem('dutySchedulerAppTitle', appTitle);
  }, [appTitle]);

  useEffect(() => {
    localStorage.setItem('dutySchedulerStartDay', startDayOfWeek.toString());
  }, [startDayOfWeek]);

  useEffect(() => {
    localStorage.setItem('dutySchedulerStartMonth', startMonth.toString());
  }, [startMonth]);

  useEffect(() => {
    try {
        localStorage.setItem('dutySchedulerPeople', JSON.stringify(people));
    } catch (e) {
        console.error("Failed to save people to localStorage", e);
    }
  }, [people]);

  // Fetch Holidays using Nager.Date API
  const fetchHolidays = useCallback(async (yr: number) => {
    setIsLoadingHolidays(true);
    setHolidayError(null);
    try {
      // Fetch for current year and next year to handle holiday period spanning year end
      const [responseCurrent, responseNext] = await Promise.all([
          fetch(`https://date.nager.at/api/v3/PublicHolidays/${yr}/US`),
          fetch(`https://date.nager.at/api/v3/PublicHolidays/${yr + 1}/US`)
      ]);
      
      if (!responseCurrent.ok) {
        throw new Error(`HTTP error! status: ${responseCurrent.status} for year ${yr}`);
      }
      if (!responseNext.ok) {
         // Don't throw, but log warning, as next year might not be needed if schedule ends early
         console.warn(`HTTP error! status: ${responseNext.status} for year ${yr + 1}. Holiday period calculation might be affected.`);
      }

      const dataCurrent: any[] = await responseCurrent.json();
      const dataNext: any[] = responseNext.ok ? await responseNext.json() : [];
      
      // Combine and format data
      const formattedHolidays: Holiday[] = [...dataCurrent, ...dataNext].map(h => ({
         date: h.date, // Assuming API returns YYYY-MM-DD string
         name: h.name 
      }));
      setHolidays(formattedHolidays);
    } catch (error: any) {
      console.error("Failed to fetch holidays:", error);
      setHolidayError(`Failed to load holidays for ${yr}. Error: ${error.message}`);
      setHolidays([]); // Set empty holidays on error
    } finally {
      setIsLoadingHolidays(false);
    }
  }, []);

  // Fetch holidays when year changes
  useEffect(() => {
    fetchHolidays(year);
  }, [year, fetchHolidays]);

  // Generate schedule when dependencies change and holidays are loaded/failed
  useEffect(() => {
    if (!isLoadingHolidays) { 
        generateSchedule(holidays);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [year, people, startDayOfWeek, startMonth, holidays, isLoadingHolidays]); 

  // Update assignment counts when schedule changes
  useEffect(() => {
    recalculateAssignments();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [schedule, holidays]); // Also depends on holidays for accurate counting

  // --- Core Logic Functions ---

  // Calculate the precise holiday period start and end dates based on user rules
  const getHolidayPeriod = (scheduleYear: number): { start: Date, end: Date } | null => {
      try {
          const christmas = new Date(Date.UTC(scheduleYear, 11, 25)); // Dec 25th of schedule year
          const newYear = new Date(Date.UTC(scheduleYear + 1, 0, 1)); // Jan 1st of next year

          const holidayPeriodStart = getMondayBefore(christmas);
          const holidayPeriodEnd = getFridayAfter(newYear);
          
          // Ensure dates are valid
          if (isNaN(holidayPeriodStart.getTime()) || isNaN(holidayPeriodEnd.getTime())) {
              throw new Error("Invalid date calculation for holiday period.");
          }

          return { start: holidayPeriodStart, end: holidayPeriodEnd };
      } catch (error) {
          console.error("Error calculating holiday period:", error);
          return null; // Return null if calculation fails
      }
  };

  // Checks if a week (defined by start/end dates) overlaps with the calculated holiday period
  const isWeekInHolidayPeriod = (weekStartDate: Date, weekEndDate: Date, scheduleYear: number): boolean => {
      const holidayPeriod = getHolidayPeriod(scheduleYear);
      if (!holidayPeriod) return false; // If period calculation failed, assume no overlap

      // Check for overlap: (WeekStart <= PeriodEnd) and (WeekEnd >= PeriodStart)
      return weekStartDate <= holidayPeriod.end && weekEndDate >= holidayPeriod.start;
  };

  const weekContainsHoliday = (startDate: Date, endDate: Date, holidayList: Holiday[]): boolean => {
      return holidayList.some(holiday => {
          // API dates are 'YYYY-MM-DD', need to compare correctly
          try {
              const holidayDate = new Date(holiday.date + 'T00:00:00Z'); // Use UTC to avoid timezone issues
              // Normalize start/end dates to midnight UTC for comparison
              const startUTC = new Date(Date.UTC(startDate.getFullYear(), startDate.getMonth(), startDate.getDate()));
              const endUTC = new Date(Date.UTC(endDate.getFullYear(), endDate.getMonth(), endDate.getDate()));
              return holidayDate >= startUTC && holidayDate <= endUTC;
          } catch (e) {
              console.error(`Invalid holiday date format: ${holiday.date}`, e);
              return false;
          }
      });
  };

  const generateSchedule = useCallback((holidayList: Holiday[]) => {
    const newSchedule: Week[] = [];
    
    // Find the first occurrence of startDayOfWeek in the startMonth of the year
    let date = new Date(Date.UTC(year, startMonth, 1)); // Use UTC
    while (date.getUTCDay() !== startDayOfWeek) {
      date.setUTCDate(date.getUTCDate() + 1);
      // Ensure we stay within the target year/month initially
      if (date.getUTCFullYear() > year || (date.getUTCFullYear() === year && date.getUTCMonth() > startMonth)) {
          // This can happen if startMonth is Dec and startDayOfWeek is early in the week
          // Reset to the first day of the start month and find the first startDayOfWeek
          date = new Date(Date.UTC(year, startMonth, 1));
          const dayDiff = (startDayOfWeek - date.getUTCDay() + 7) % 7;
          date.setUTCDate(date.getUTCDate() + dayDiff);
          break;
      }
    }

    // Loop while the week's start date is within the selected calendar year
    while (date.getUTCFullYear() === year) {
        const startDate = new Date(date); // Clone date
        const endDate = new Date(date);   // Clone date
        endDate.setUTCDate(endDate.getUTCDate() + 6);

        // *** NEW: Check if this week falls within the precise holiday period ***
        const isMarkedAsHolidayPeriod = isWeekInHolidayPeriod(startDate, endDate, year);
        
        const weekObj: Week = {
            startDate: startDate,
            endDate: endDate,
            // Mark as "Holiday Period" if it falls within the calculated range
            assignedTo: isMarkedAsHolidayPeriod ? "Holiday Period" : null, 
            hasHoliday: weekContainsHoliday(startDate, endDate, holidayList),
            isHolidayPeriod: isMarkedAsHolidayPeriod,
            notes: "" // Initialize notes field
        };

        newSchedule.push(weekObj);

        // Move to the next week's start day
        date.setUTCDate(date.getUTCDate() + 7);
    }

    assignPeopleToWeeks(newSchedule, holidayList);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [year, startMonth, startDayOfWeek, people]); // Dependencies for schedule generation structure


  const assignPeopleToWeeks = (currentSchedule: Week[], holidayList: Holiday[]) => {
    const counts: HolidayCounts = {};
    people.forEach((person: Person) => {
      counts[person.id] = { total: 0, holiday: 0, lastAssigned: -Infinity };
    });

    const isPersonOnLeave = (personId: number, startDate: Date, endDate: Date): boolean => {
      const person = people.find((p: { id: number; }) => p.id === personId);
      if (!person || !person.leave) return false;
      return person.leave.some((leave: { start: string; end: string; }) => {
        try {
            const leaveStart = new Date(leave.start + 'T00:00:00Z');
            const leaveEnd = new Date(leave.end + 'T00:00:00Z');
            // Normalize week dates to UTC midnight
            const startUTC = new Date(Date.UTC(startDate.getFullYear(), startDate.getMonth(), startDate.getDate()));
            const endUTC = new Date(Date.UTC(endDate.getFullYear(), endDate.getMonth(), endDate.getDate()));
            // Check if leave period [leaveStart, leaveEnd] overlaps with week [startUTC, endUTC]
            return leaveStart <= endUTC && leaveEnd >= startUTC;
        } catch (e) {
            console.error(`Invalid leave date format for person ${personId}: ${leave.start} or ${leave.end}`, e);
            return false;
        }
      });
    };

    const updatedSchedule = currentSchedule.map((week, index) => {
        // *** If the week is marked as a holiday period, skip assignment ***
        if (week.isHolidayPeriod) return week; 

        const availablePeople = people.filter((person: { id: number; }) =>
            !isPersonOnLeave(person.id, week.startDate, week.endDate)
        );

        if (availablePeople.length === 0) {
            return { ...week, assignedTo: "No one available" };
        }

        // Sort available people
        availablePeople.sort((a: Person, b: Person) => {
            const countA = counts[a.id];
            const countB = counts[b.id];
            if (countA.total !== countB.total) return countA.total - countB.total;
            if (countA.holiday !== countB.holiday) return countA.holiday - countB.holiday;
            return (countA.lastAssigned ?? -Infinity) - (countB.lastAssigned ?? -Infinity);
        });

        const assignedPerson = availablePeople[0];
        const assignedPersonId = assignedPerson.id;

        // Update counts for the assigned person
        counts[assignedPersonId].total += 1;
        counts[assignedPersonId].lastAssigned = index;

        // Recalculate hasHoliday based on the fetched list for accurate counting
        const currentWeekHasHoliday = weekContainsHoliday(week.startDate, week.endDate, holidayList);
        if (currentWeekHasHoliday) {
            counts[assignedPersonId].holiday += 1;
        }

        // Return the updated week object
        return { 
            ...week, 
            assignedTo: assignedPersonId,
            hasHoliday: currentWeekHasHoliday // Ensure hasHoliday is correctly set
        };
    });

    setSchedule(updatedSchedule);
    // Remove lastAssigned before setting final counts state
    const finalCounts: HolidayCounts = {};
    Object.keys(counts).forEach(key => {
        const personId = parseInt(key, 10);
        finalCounts[personId] = { total: counts[personId].total, holiday: counts[personId].holiday };
    });
    setHolidayCounts(finalCounts); 
  };

  // Recalculate assignment counts (used by useEffect)
  const recalculateAssignments = () => {
    const counts: HolidayCounts = {};
    people.forEach((person: Person) => {
      counts[person.id] = { total: 0, holiday: 0 };
    });

    schedule.forEach((week: { assignedTo: any; hasHoliday: any; }) => {
      if (typeof week.assignedTo === 'number') {
        const personId = week.assignedTo;
        if (counts[personId]) {
          counts[personId].total += 1;
          // Use the existing week.hasHoliday flag which should be correct after assignment
          if (week.hasHoliday) {
            counts[personId].holiday += 1;
          }
        }
      }
    });
    setHolidayCounts(counts);
  };

  // --- Event Handlers & UI Logic ---

  const handleTitleSave = () => {
    const trimmedTitle = tempTitle.trim();
    if (trimmedTitle) {
        setAppTitle(trimmedTitle);
        localStorage.setItem('dutySchedulerAppTitle', trimmedTitle);
    }
    setIsEditingTitle(false);
  };

  const handleTitleCancel = () => {
    setTempTitle(appTitle); // Reset temp title
    setIsEditingTitle(false);
  };

  const addPerson = () => {
    const trimmedName = newPersonName.trim();
    if (trimmedName && people.length < MAX_PEOPLE) {
      const newId = people.length > 0 ? Math.max(...people.map((p: { id: any; }) => p.id)) + 1 : 1;
      setPeople([...people, { id: newId, name: trimmedName, leave: [] }]);
      setNewPersonName("");
    }
  };

  const removePerson = (idToRemove: number) => {
    // Also remove assignments for this person in the current schedule
    const newSchedule = schedule.map((week: Week) => {
        if (week.assignedTo === idToRemove) {
            // Re-run assignment logic for this week potentially?
            // Simpler: just unassign for now. Re-generation handles re-assignment.
            return { ...week, assignedTo: null }; 
        }
        return week;
    });
    setSchedule(newSchedule); // Update schedule first
    setPeople(people.filter((person: Person) => person.id !== idToRemove)); // Then remove person
  };
  
  const updatePersonName = (idToUpdate: number, newName: string) => {
     setPeople(people.map((p: Person) => p.id === idToUpdate ? {...p, name: newName.trim() } : p));
  };

  const addLeave = () => {
    if (selectedPersonIdForLeave !== null && leaveDate && leaveEndDate) {
      try {
          const start = new Date(leaveDate + 'T00:00:00Z');
          const end = new Date(leaveEndDate + 'T00:00:00Z');
          if (start > end) {
            alert("Leave end date must be on or after the start date.");
            return;
          }
          
          setPeople(people.map((person: Person) => {
            if (person.id === selectedPersonIdForLeave) {
              const currentLeave = person.leave || [];
              // Avoid adding duplicate leave periods
              const alreadyExists = currentLeave.some((l: { start: any; end: any; }) => l.start === leaveDate && l.end === leaveEndDate);
              if (alreadyExists) {
                  alert("This leave period already exists for this person.");
                  return person;
              }
              return {
                ...person,
                leave: [...currentLeave, { start: leaveDate, end: leaveEndDate }]
              };
            }
            return person;
          }));
          
          // Reset form & close
          setLeaveDate("");
          setLeaveEndDate("");
          setSelectedPersonIdForLeave(null);
          setShowLeaveForm(false);
          // Schedule regeneration is handled by useEffect watching 'people'
      } catch (e) {
          alert("Invalid date format for leave.");
          console.error("Leave date error:", e);
      }
    }
  };

  const removeLeave = (personId: number, leaveIndex: number) => {
    setPeople(people.map((person: Person) => {
      if (person.id === personId) {
        const newLeave = [...(person.leave || [])];
        if (leaveIndex >= 0 && leaveIndex < newLeave.length) {
            newLeave.splice(leaveIndex, 1);
        }
        return { ...person, leave: newLeave };
      }
      return person;
    }));
     // Schedule regeneration is handled by useEffect watching 'people'
  };

  const handleWeekAssignmentChange = (index: number, personIdStr: string) => {
    const newSchedule = [...schedule];
    if (index >= 0 && index < newSchedule.length) {
        const personId = personIdStr ? parseInt(personIdStr, 10) : null;
        // Prevent assignment if it's a holiday period week
        if (!newSchedule[index].isHolidayPeriod) {
            newSchedule[index].assignedTo = personId;
            setSchedule(newSchedule);
            // Recalculate counts immediately for responsiveness
            recalculateAssignments(); 
        } else {
            alert("Cannot assign duty during the designated holiday period.");
            // Optionally reset the dropdown visually if needed, though state didn't change
        }
    }
  };

  const openNotesModal = (index: number) => {
    if (index >= 0 && index < schedule.length) {
        setCurrentWeekIndexForNotes(index);
        setCurrentNotes(schedule[index].notes || "");
        setShowNotesModal(true);
    }
  };

  const saveNotes = () => {
    if (currentWeekIndexForNotes !== null && currentWeekIndexForNotes >= 0 && currentWeekIndexForNotes < schedule.length) {
      const newSchedule = [...schedule];
      newSchedule[currentWeekIndexForNotes].notes = currentNotes.trim();
      setSchedule(newSchedule);
    }
    closeNotesModal(); // Close and reset state
  };

  const closeNotesModal = () => {
    setShowNotesModal(false);
    setCurrentWeekIndexForNotes(null);
    setCurrentNotes("");
  };

  const formatDate = (date: Date | string | undefined): string => {
      if (!date) return '';
      try {
          // Handle both Date objects and date strings (e.g., from import)
          const d = date instanceof Date ? date : new Date(date);
          // Check if the date is valid
          if (isNaN(d.getTime())) return 'Invalid Date';
          // Format using UTC date parts to avoid timezone shifts in display
          return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', timeZone: 'UTC' });
      } catch (e) {
          console.error("Error formatting date:", date, e);
          return 'Invalid Date';
      }
  };

  // Helper to format date as YYYY-MM-DD for CSV
  const formatDateYYYYMMDD = (date: Date | string | undefined): string => {
      if (!date) return '';
      try {
          const d = date instanceof Date ? date : new Date(date);
          if (isNaN(d.getTime())) return 'Invalid Date';
          const year = d.getUTCFullYear();
          const month = (d.getUTCMonth() + 1).toString().padStart(2, '0');
          const day = d.getUTCDate().toString().padStart(2, '0');
          return `${year}-${month}-${day}`;
      } catch (e) {
          console.error("Error formatting date for CSV:", date, e);
          return 'Invalid Date';
      }
  };

  const getPersonName = (id: number | string | null): string => {
    if (typeof id === 'number') {
        const person = people.find((p: { id: number; }) => p.id === id);
        return person ? person.name : "Unknown Person";
    }
    if (typeof id === 'string') {
        return id; // e.g., "Holiday Period", "No one available"
    }
    return "Unassigned";
  };

  const getPersonIdFromName = (name: string): number | string | null => {
      const person = people.find((p: { name: string; }) => p.name.toLowerCase() === name.toLowerCase());
      if (person) return person.id;
      // Handle special status strings
      if (name === "Holiday Period" || name === "No one available" || name === "Unassigned" || name === "") return name || null;
      console.warn(`Could not find person ID for name: ${name}`);
      return null; // Or handle as error / unassigned
  };

  const getHolidaysInWeek = (startDate: Date | string | undefined, endDate: Date | string | undefined): Holiday[] => {
    if (!holidays || holidays.length === 0 || !startDate || !endDate) return [];
    try {
      const start = startDate instanceof Date ? startDate : new Date(startDate);
      const end = endDate instanceof Date ? endDate : new Date(endDate);
      if (isNaN(start.getTime()) || isNaN(end.getTime())) return [];
      
      return holidays.filter((holiday: { date: string; }) => {
          try {
              const holidayDate = new Date(holiday.date + 'T00:00:00Z');
              const startUTC = new Date(Date.UTC(start.getFullYear(), start.getMonth(), start.getDate()));
              const endUTC = new Date(Date.UTC(end.getFullYear(), end.getMonth(), end.getDate()));
              return holidayDate >= startUTC && holidayDate <= endUTC;
          } catch { return false; }
      });
    } catch {
      return [];
    }
  };

  // --- CSV Import / Export --- 

  // Helper to escape CSV fields containing commas, quotes, or newlines
  const escapeCSV = (field: string | number | null | undefined): string => {
      if (field === null || field === undefined) return '';
      let str = String(field);
      // If the string contains a comma, newline, or double quote, enclose it in double quotes
      if (str.includes(',') || str.includes('\n') || str.includes('"')) {
          // Escape existing double quotes by doubling them
          str = str.replace(/"/g, '""');
          // Enclose the entire field in double quotes
          str = `"${str}"`;
      }
      return str;
  };

  // Helper to unescape CSV fields
  const unescapeCSV = (field: string): string => {
      if (field.startsWith('"') && field.endsWith('"')) {
          // Remove surrounding quotes
          let str = field.slice(1, -1);
          // Unescape doubled double quotes
          str = str.replace(/""/g, '"');
          return str;
      }
      return field;
  };

  const exportScheduleCSV = () => {
    if (schedule.length === 0) {
        alert("No schedule data to export.");
        return;
    }

    // Map schedule data to CSV rows
    const rows = schedule.map((week: Week) => {
        const startDateStr = formatDateYYYYMMDD(week.startDate);
        const endDateStr = formatDateYYYYMMDD(week.endDate);
        const assignedName = getPersonName(week.assignedTo);
        const holidaysStr = getHolidaysInWeek(week.startDate, week.endDate)
                              .map(h => `${h.date}: ${h.name}`)
                              .join('\n'); // Use newline within the cell for multiple holidays
        const notesStr = week.notes || '';

        return [
            escapeCSV(startDateStr),
            escapeCSV(endDateStr),
            escapeCSV(assignedName),
            escapeCSV(holidaysStr),
            escapeCSV(notesStr)
        ].join(',');
    });

    // Combine headers and rows
    const csvContent = [CSV_HEADERS.join(','), ...rows].join('\n');

    try {
        // Create Blob and trigger download
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `duty-schedule-${year}-${appTitle.replace(/\s+/g, '_')}.csv`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    } catch (e) {
        console.error("CSV Export failed:", e);
        alert("Failed to export schedule as CSV.");
    }
  };

  const importScheduleCSV = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const csvContent = event.target?.result as string;
        if (!csvContent) throw new Error("File is empty or could not be read.");

        // Basic CSV parsing (split lines, then commas, handling quoted fields)
        const lines = csvContent.split(/\r?\n/).filter(line => line.trim() !== ''); // Split lines and remove empty ones
        if (lines.length < 2) throw new Error("CSV file must contain headers and at least one data row.");

        // Parse header (assuming simple comma separation for header)
        const headerLine = lines[0];
        const headers = headerLine.split(',').map(h => h.trim());

        // Validate headers (basic check)
        if (headers.length !== CSV_HEADERS.length || !CSV_HEADERS.every((h, i) => headers[i] === h)) {
            console.error("Expected Headers:", CSV_HEADERS);
            console.error("Found Headers:", headers);
            throw new Error(`Invalid CSV headers. Expected: ${CSV_HEADERS.join(', ')}`);
        }

        // Find column indices based on headers
        const colIndices: { [key: string]: number } = {};
        CSV_HEADERS.forEach(header => {
            const index = headers.indexOf(header);
            if (index === -1) throw new Error(`Missing required header: ${header}`);
            colIndices[header] = index;
        });

        // Parse data rows
        const importedSchedule: Week[] = [];
        const dataLines = lines.slice(1);

        for (let i = 0; i < dataLines.length; i++) {
            const line = dataLines[i];
            // More robust CSV parsing needed to handle quoted commas/newlines
            // This is a simplified parser - consider a library for complex CSVs
            const values: string[] = [];
            let currentVal = '';
            let inQuotes = false;
            for (let char of line) {
                if (char === '"') {
                    inQuotes = !inQuotes;
                } else if (char === ',' && !inQuotes) {
                    values.push(unescapeCSV(currentVal.trim()));
                    currentVal = '';
                } else {
                    currentVal += char;
                }
            }
            values.push(unescapeCSV(currentVal.trim())); // Add the last value

            if (values.length !== CSV_HEADERS.length) {
                console.warn(`Skipping row ${i + 1}: Incorrect number of columns. Expected ${CSV_HEADERS.length}, got ${values.length}. Line: ${line}`);
                continue;
            }

            try {
                const startDateStr = values[colIndices["Week Start Date"]];
                const endDateStr = values[colIndices["Week End Date"]];
                const assignedName = values[colIndices["Assigned To"]];
                // const holidaysStr = values[colIndices["Holidays"]]; // Holidays are derived, not imported directly
                const notes = values[colIndices["Notes"]];

                const startDate = new Date(startDateStr + 'T00:00:00Z');
                const endDate = new Date(endDateStr + 'T00:00:00Z');

                if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
                    console.warn(`Skipping row ${i + 1}: Invalid date format. Start: ${startDateStr}, End: ${endDateStr}`);
                    continue;
                }

                // Attempt to find person ID from name
                const assignedToValue = getPersonIdFromName(assignedName);

                // Re-calculate holiday status based on imported dates and fetched holidays
                // Use the year from the imported start date for holiday period check
                const isHolidayPeriod = isWeekInHolidayPeriod(startDate, endDate, startDate.getUTCFullYear()); 
                const hasHoliday = weekContainsHoliday(startDate, endDate, holidays); // Use current holidays state

                importedSchedule.push({
                    startDate,
                    endDate,
                    assignedTo: isHolidayPeriod ? "Holiday Period" : assignedToValue,
                    hasHoliday,
                    isHolidayPeriod,
                    notes
                });
            } catch (parseError: any) {
                console.warn(`Skipping row ${i + 1} due to error: ${parseError.message}. Line: ${line}`);
            }
        }

        if (importedSchedule.length === 0) {
            throw new Error("No valid schedule data could be imported from the CSV.");
        }

        // Update the schedule state
        setSchedule(importedSchedule);
        // Optional: Update year/month based on imported data? For now, keep UI settings.
        // Recalculate assignments based on the imported schedule
        recalculateAssignments(); 

        alert(`Successfully imported ${importedSchedule.length} weeks from CSV. Assignment counts updated.`);

      } catch (error: any) {
        console.error("CSV Import error:", error);
        alert(`Error importing schedule from CSV: ${error.message}. Please check file format and headers.`);
      }
    };
    reader.onerror = () => {
        alert("Failed to read the file.");
    };
    reader.readAsText(file);
    // Clear the input value to allow importing the same file again
    if (e.target) e.target.value = ''; 
  };

  // --- Render --- 
  return (
    <div className="flex flex-col p-4 max-w-full font-sans">
      {/* Header: Title, Year, Settings, Import/Export - INLINE LAYOUT */}
      <div className="flex flex-wrap items-center mb-6 gap-x-4 gap-y-2 border-b pb-4">
        {/* Title and Edit Button */}
        <div className="flex items-center gap-2 flex-shrink-0 mr-4">
          {isEditingTitle ? (
            <div className="flex items-center gap-1">
              <input 
                type="text"
                value={tempTitle}
                onChange={(e: { target: { value: any; }; }) => setTempTitle(e.target.value)}
                className="text-xl font-bold border-b-2 border-blue-500 focus:outline-none px-1"
                autoFocus
                onKeyDown={(e: { key: string; }) => e.key === 'Enter' && handleTitleSave()}
                maxLength={50}
              />
              <button onClick={handleTitleSave} className="text-green-600 hover:text-green-800 p-1" title="Save Title"><Save size={18} /></button>
              <button onClick={handleTitleCancel} className="text-red-600 hover:text-red-800 p-1" title="Cancel Edit"><X size={18} /></button>
            </div>
          ) : (
            <div className="flex items-center gap-1">
              <h1 className="text-xl font-bold" title={appTitle}>{appTitle.length > 30 ? appTitle.substring(0, 27) + '...' : appTitle}</h1>
              <button onClick={() => setIsEditingTitle(true)} className="text-gray-500 hover:text-blue-600 p-1" title="Edit Title"><Edit size={16} /></button>
            </div>
          )}
        </div>

        {/* Settings Dropdowns */}
        <div className="flex items-center">
          <label htmlFor="year-select" className="mr-1 font-medium text-xs">Year:</label>
          <select
            id="year-select"
            className="border rounded px-2 py-0.5 text-xs bg-white focus:outline-none focus:ring-1 focus:ring-blue-500"
            value={year}
            onChange={(e: { target: { value: string; }; }) => setYear(parseInt(e.target.value, 10))}
          >
            {Array.from({ length: SCHEDULE_YEAR_RANGE }, (_, i) => currentYear - Math.floor(SCHEDULE_YEAR_RANGE / 2) + i).map((y) => (
              <option key={y} value={y}>{y}</option>
            ))}
          </select>
        </div>
        <div className="flex items-center">
          <label htmlFor="start-month-select" className="mr-1 font-medium text-xs">Start Month:</label>
          <select
            id="start-month-select"
            className="border rounded px-2 py-0.5 text-xs bg-white focus:outline-none focus:ring-1 focus:ring-blue-500"
            value={startMonth}
            onChange={(e: { target: { value: string; }; }) => setStartMonth(parseInt(e.target.value, 10))}
          >
            {monthNames.map((name, index) => (
              <option key={index} value={index}>{name}</option>
            ))}
          </select>
        </div>
        <div className="flex items-center">
          <label htmlFor="start-day-select" className="mr-1 font-medium text-xs">Week Starts:</label>
          <select
            id="start-day-select"
            className="border rounded px-2 py-0.5 text-xs bg-white focus:outline-none focus:ring-1 focus:ring-blue-500"
            value={startDayOfWeek}
            onChange={(e: { target: { value: string; }; }) => setStartDayOfWeek(parseInt(e.target.value, 10))}
          >
            {dayNames.map((name, index) => (
              <option key={index} value={index}>{name}</option>
            ))}
          </select>
        </div>

        {/* Import/Export Buttons */}
        <div className="flex items-center gap-2">
          <button
            className="flex items-center bg-teal-500 text-white px-2 py-1 rounded hover:bg-teal-600 text-xs shadow-sm"
            onClick={exportScheduleCSV}
            title="Export current schedule to a CSV file"
          >
            <FileText size={14} className="mr-1" /> Export CSV
          </button>
          <div className="relative">
            <input
              type="file"
              id="import-csv-file"
              className="hidden"
              accept=".csv"
              onChange={importScheduleCSV} // Changed to CSV import handler
            />
            <label
              htmlFor="import-csv-file"
              className="flex items-center bg-green-500 text-white px-2 py-1 rounded hover:bg-green-600 cursor-pointer text-xs shadow-sm"
              title="Import schedule from a CSV file"
            >
              <Upload size={14} className="mr-1" /> Import CSV
            </label>
          </div>
        </div>
      </div>
      {/* End Inline Header Layout */}
      
      {/* Holiday Loading/Error Status */}
      {isLoadingHolidays && <div className="text-center text-blue-600 mb-4 p-2 bg-blue-50 rounded">Loading holidays...</div>}
      {holidayError && <div className="text-center text-red-600 mb-4 bg-red-100 border border-red-400 rounded p-2 text-sm">{holidayError}</div>}

      {/* Main Content: People, Schedule, Analytics */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Left column: People Management */}
        <div className="border rounded p-4 bg-gray-50 shadow-sm">
          <div className="flex justify-between items-center mb-4">
            <h2 className="text-lg font-medium">Team Members</h2>
            <span className="text-sm text-gray-500">{people.length}/{MAX_PEOPLE}</span>
          </div>

          {/* Add new person form */}
          <div className="flex mb-4">
            <input
              type="text"
              className="flex-grow border rounded-l px-3 py-2 text-sm focus:outline-none focus:ring-1 focus:ring-blue-500"
              placeholder="Add new person name"
              value={newPersonName}
              onChange={(e: { target: { value: any; }; }) => setNewPersonName(e.target.value)}
              maxLength={30}
              onKeyDown={(e: { key: string; }) => e.key === 'Enter' && addPerson()}
            />
            <button
              className="bg-blue-500 text-white px-3 py-2 rounded-r hover:bg-blue-600 disabled:bg-gray-400 disabled:cursor-not-allowed"
              onClick={addPerson}
              disabled={!newPersonName.trim() || people.length >= MAX_PEOPLE}
              title={people.length >= MAX_PEOPLE ? `Maximum ${MAX_PEOPLE} people allowed` : "Add Person"}
            >
              <Plus size={18} />
            </button>
          </div>

          {/* People list */}
          <div className="space-y-2 mt-2 max-h-96 overflow-y-auto pr-2">
            {people.map((person: { id: number; name: any; leave: any[]; }) => (
              <div key={person.id} className="border rounded p-3 bg-white shadow-sm">
                <div className="flex justify-between items-center mb-2">
                   <input 
                    type="text" 
                    className="font-medium border-b border-dashed border-gray-400 focus:border-blue-500 focus:outline-none bg-transparent text-sm w-2/3"
                    value={person.name}
                    onChange={(e: { target: { value: string; }; }) => updatePersonName(person.id, e.target.value)}
                    maxLength={30}
                  />
                  <div className="flex items-center space-x-1">
                    <button
                      className="text-blue-500 hover:text-blue-700 text-xs font-medium p-1 rounded hover:bg-blue-100"
                      onClick={() => {
                        setSelectedPersonIdForLeave(person.id);
                        setShowLeaveForm(true);
                      }}
                      title={`Add Leave for ${person.name}`}
                    >
                      Add Leave
                    </button>
                    <button
                      className="text-red-500 hover:text-red-700 p-1 rounded hover:bg-red-100"
                      onClick={() => removePerson(person.id)}
                      title={`Remove ${person.name}`}
                    >
                      <Trash2 size={14} />
                    </button>
                  </div>
                </div>

                {/* Leave periods */}
                {person.leave && person.leave.length > 0 && (
                  <div className="mt-2 border-t pt-2">
                    <p className="text-xs text-gray-600 mb-1 font-medium">Leave Periods:</p>
                    <ul className="space-y-1">
                      {person.leave.map((leave: { start: string | Date | undefined; end: string | Date | undefined; }, index: number) => (
                        <li key={index} className="flex justify-between items-center text-xs text-gray-700">
                          <span>
                            {formatDate(leave.start)} - {formatDate(leave.end)}
                          </span>
                          <button
                            className="text-red-500 hover:text-red-700 p-0.5 rounded hover:bg-red-100"
                            onClick={() => removeLeave(person.id, index)}
                            title="Remove Leave Period"
                          >
                            <Trash2 size={12} />
                          </button>
                        </li>
                      ))}
                    </ul>
                  </div>
                )}
              </div>
            ))}
             {people.length === 0 && <p className="text-sm text-gray-500 text-center py-4">No team members added yet.</p>}
          </div>
        </div>

        {/* Center column: Schedule Table */}
        <div className="border rounded p-4 lg:col-span-2 bg-white shadow-sm">
          <h2 className="text-lg font-medium mb-4">Duty Schedule for {year}</h2>
          <div className="overflow-x-auto">
            <table className="min-w-full border-collapse text-sm">
              <thead>
                <tr className="bg-gray-100">
                  <th className="border px-3 py-2 text-left font-semibold text-gray-700">Week ({dayNames[startDayOfWeek]} - {dayNames[(startDayOfWeek + 6) % 7]})</th>
                  <th className="border px-3 py-2 text-left font-semibold text-gray-700">Assigned To</th>
                  <th className="border px-3 py-2 text-left font-semibold text-gray-700">Holidays</th>
                  <th className="border px-3 py-2 text-center font-semibold text-gray-700">Notes</th>
                </tr>
              </thead>
              <tbody>
                {schedule.length > 0 ? schedule.map((week: { startDate: string | Date | undefined; isHolidayPeriod: any; hasHoliday: any; endDate: string | Date | undefined; assignedTo: string | number | null; notes: any; }, index: number) => (
                  <tr key={`${formatDate(week.startDate)}-${index}`} className={`${week.isHolidayPeriod ? "bg-red-50" : week.hasHoliday ? "bg-blue-50" : "hover:bg-gray-50"} transition-colors duration-150`}>
                    <td className="border px-3 py-2 whitespace-nowrap">
                      {formatDate(week.startDate)} - {formatDate(week.endDate)}
                    </td>
                    <td className={`border px-3 py-2 ${week.isHolidayPeriod ? "font-semibold text-red-600" : ""}`}>
                      {week.assignedTo === "Holiday Period" 
                        ? "Holiday Period"
                        : week.assignedTo === "No one available" 
                        ? <span className="text-orange-600 font-medium">No one available</span>
                        : (
                          <select
                            className="w-full bg-transparent border-b border-dashed border-gray-400 focus:border-blue-500 focus:outline-none text-sm py-1"
                            value={typeof week.assignedTo === 'number' ? week.assignedTo : ""}
                            onChange={(e: { target: { value: string; }; }) => handleWeekAssignmentChange(index, e.target.value)}
                            title={`Current: ${getPersonName(week.assignedTo)}`}
                            // Disable assignment dropdown if it's a holiday period week
                            disabled={week.isHolidayPeriod} 
                          >
                            <option value="">Unassigned</option>
                            {people.map((person: { id: any; name: any; }) => (
                              <option key={person.id} value={person.id}>
                                {person.name}
                              </option>
                            ))}
                          </select>
                        )}
                    </td>
                    <td className="border px-3 py-2 text-xs">
                      {getHolidaysInWeek(week.startDate, week.endDate).map((holiday, i) => (
                        <div key={i} title={holiday.name}>
                          {formatDate(new Date(holiday.date))}: {holiday.name.length > 25 ? holiday.name.substring(0, 22) + '...' : holiday.name}
                        </div>
                      ))}
                      {getHolidaysInWeek(week.startDate, week.endDate).length === 0 && <span className="text-gray-400">-</span>}
                    </td>
                     <td className="border px-3 py-2 text-center">
                        <button 
                            onClick={() => openNotesModal(index)}
                            className={`p-1 rounded ${week.notes ? 'text-yellow-600 hover:bg-yellow-100' : 'text-gray-400 hover:text-gray-600 hover:bg-gray-100'}`}
                            title={week.notes ? "View/Edit Note" : "Add Note"}
                        >
                            <StickyNote size={16} />
                        </button>
                    </td>
                  </tr>
                )) : (
                    <tr>
                        {/* Use numeric colSpan */} 
                        <td colSpan={4} className="text-center py-8 text-gray-500 border">
                            {isLoadingHolidays ? "Loading schedule..." : "No schedule generated. Check settings or add team members."}
                        </td>
                    </tr>
                )}
              </tbody>
            </table>
          </div>
        </div>
      </div>

      {/* Analytics Section */}
      <div className="mt-6 border rounded p-4 bg-white shadow-sm">
        <h2 className="text-lg font-medium mb-4">Assignment Distribution ({year})</h2>
        <div className="overflow-x-auto">
          <table className="min-w-full border-collapse text-sm">
            <thead>
              <tr className="bg-gray-100">
                <th className="border px-3 py-2 text-left font-semibold text-gray-700">Person</th>
                <th className="border px-3 py-2 text-center font-semibold text-gray-700">Total Duties</th>
                <th className="border px-3 py-2 text-center font-semibold text-gray-700">Holiday Duties</th>
              </tr>
            </thead>
            <tbody>
              {people.length > 0 ? people.map((person: Person) => {
                const counts = holidayCounts[person.id] || { total: 0, holiday: 0 };
                return (
                  <tr key={person.id} className="hover:bg-gray-50">
                    <td className="border px-3 py-2">{person.name}</td>
                    <td className="border px-3 py-2 text-center">{counts.total}</td>
                    <td className="border px-3 py-2 text-center">{counts.holiday}</td>
                  </tr>
                );
              }) : (
                 <tr>
                    {/* Use numeric colSpan */} 
                    <td colSpan={3} className="text-center py-4 text-gray-500 border">No team members to analyze.</td>
                 </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Leave Form Modal */}
      {showLeaveForm && (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-lg p-6 w-full max-w-md shadow-xl">
            <div className="flex justify-between items-center mb-4">
                <h3 className="text-lg font-medium">Add Leave Period</h3>
                <button onClick={() => setShowLeaveForm(false)} className="text-gray-500 hover:text-gray-800"><X size={20}/></button>
            </div>
            <div className="space-y-4">
              <div>
                <label htmlFor="leave-person-select" className="block text-sm font-medium mb-1">Person</label>
                <select
                  id="leave-person-select"
                  className="w-full border rounded px-3 py-2 text-sm bg-white focus:outline-none focus:ring-1 focus:ring-blue-500"
                  value={selectedPersonIdForLeave ?? ""} // Handle null case
                  onChange={(e: { target: { value: string; }; }) => setSelectedPersonIdForLeave(e.target.value ? parseInt(e.target.value, 10) : null)}
                >
                  <option value="" disabled>Select a person</option>
                  {people.map((person: { id: any; name: any; }) => (
                    <option key={person.id} value={person.id}>
                      {person.name}
                    </option>
                  ))}
                </select>
              </div>
              <div>
                <label htmlFor="leave-start-date" className="block text-sm font-medium mb-1">Start Date</label>
                <input
                  id="leave-start-date"
                  type="date"
                  className="w-full border rounded px-3 py-2 text-sm focus:outline-none focus:ring-1 focus:ring-blue-500"
                  value={leaveDate}
                  onChange={(e: { target: { value: any; }; }) => setLeaveDate(e.target.value)}
                />
              </div>
              <div>
                <label htmlFor="leave-end-date" className="block text-sm font-medium mb-1">End Date</label>
                <input
                  id="leave-end-date"
                  type="date"
                  className="w-full border rounded px-3 py-2 text-sm focus:outline-none focus:ring-1 focus:ring-blue-500"
                  value={leaveEndDate}
                  onChange={(e: { target: { value: any; }; }) => setLeaveEndDate(e.target.value)}
                  min={leaveDate} // Prevent end date being before start date
                />
              </div>
              <div className="flex justify-end space-x-2 mt-6">
                <button
                  className="border rounded px-4 py-2 text-sm hover:bg-gray-100"
                  onClick={() => {
                    setShowLeaveForm(false);
                    setSelectedPersonIdForLeave(null);
                    setLeaveDate("");
                    setLeaveEndDate("");
                  }}
                >
                  Cancel
                </button>
                <button
                  className="bg-blue-500 text-white rounded px-4 py-2 text-sm hover:bg-blue-600 disabled:bg-gray-400 disabled:cursor-not-allowed"
                  onClick={addLeave}
                  disabled={selectedPersonIdForLeave === null || !leaveDate || !leaveEndDate}
                >
                  Add Leave
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
      
      {/* Notes Modal */}
       {showNotesModal && currentWeekIndexForNotes !== null && (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-lg p-6 w-full max-w-lg shadow-xl">
             <div className="flex justify-between items-center mb-4">
                <h3 className="text-lg font-medium">
                    Week Notes ({formatDate(schedule[currentWeekIndexForNotes]?.startDate)} - {formatDate(schedule[currentWeekIndexForNotes]?.endDate)})
                </h3>
                <button onClick={closeNotesModal} className="text-gray-500 hover:text-gray-800"><X size={20}/></button>
            </div>
            <textarea
              className="w-full border rounded px-3 py-2 text-sm h-40 resize-y focus:outline-none focus:ring-1 focus:ring-blue-500"
              value={currentNotes}
              onChange={(e: { target: { value: any; }; }) => setCurrentNotes(e.target.value)}
              placeholder="Add notes for this week..."
              maxLength={500} // Add max length for notes
            />
            <div className="flex justify-end space-x-2 mt-4">
              <button
                className="border rounded px-4 py-2 text-sm hover:bg-gray-100"
                onClick={closeNotesModal}
              >
                Cancel
              </button>
              <button
                className="bg-yellow-500 text-white rounded px-4 py-2 text-sm hover:bg-yellow-600"
                onClick={saveNotes}
              >
                Save Notes
              </button>
            </div>
          </div>
        </div>
      )}

    </div> // End main container
  );
};

export default DutyRotationScheduler;


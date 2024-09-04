import { Calendar } from '@bryntum/calendar';
import '@bryntum/calendar/calendar.stockholm.css';

const calendar = new Calendar({
    appendTo  : 'calendar',
    mode      : 'month',
    timeZone  : 'UTC',
    date      : new Date(2024, 8, 1),
    resources : [
        {
            id         : 1,
            name       : 'Default Calendar',
            eventColor : '#217346'
        }
    ],

    events : [
        {
            id         : 1,
            name       : 'Meeting',
            startDate  : '2024-09-23T10:00:00',
            endDate    : '2024-09-23T11:00:00',
            resourceId : 1
        }
    ]
});
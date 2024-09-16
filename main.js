import { Calendar, StringHelper } from '@bryntum/calendar';
import '@bryntum/calendar/calendar.stockholm.css';
import { signIn } from './auth.js';
import { createTask, deleteTask, getTasks, updateTask } from './graph.js';
import CustomEventModel from './lib/CustomEventModel.js';

const signInLink = document.getElementById('signin');

const calendar = new Calendar({
    appendTo   : 'calendar',
    mode       : 'month',
    timeZone   : 'UTC',
    date       : new Date(2024, 8, 1),
    eventStore : {
        modelClass : CustomEventModel
    },
    resources : [
        {
            id         : 1,
            name       : 'Default Calendar',
            eventColor : 'green'
        }
    ],
    features : {
        eventEdit : {
            items : {
                nameField : {
                    required : true
                },
                // Custom fields
                percentCompleteField : {
                    type        : 'combo',
                    label       : 'Progress',
                    name        : 'percentComplete',
                    multiSelect : false,
                    required    : true,
                    items       : [
                        {
                            value : 0,
                            text  : 'Not started'
                        },
                        {
                            value : 50,
                            text  : 'In progress'
                        },
                        { value : 100, text : 'Completed' }
                    ]
                },
                descriptionField : {
                    type  : 'textarea',
                    label : 'Notes',
                    // Name of the field in the event record to read/write data to
                    // NOTE: Make sure your EventModel has this field for this to link up correctly
                    name  : 'description'
                }
            }
        }
    },
    modes : {
        month : {
            // Render an icon showing progress state (editable in the event editor)
            eventRenderer : ({ eventRecord, renderData }) => {
                if (eventRecord.percentComplete === 0) {
                    renderData.eventColor = '#605e5c';
                }
                if (eventRecord.percentComplete === 50) {
                    renderData.eventColor = '#327eaa';
                    renderData.iconCls['b-fa b-fa-hourglass-half'] = 1;
                }
                if (eventRecord.percentComplete === 100) {
                    renderData.eventColor = '#107c41';
                    renderData.iconCls['b-fa b-fa-exclamation'] = 1;
                }
                return `
              <span class="b-event-name">${StringHelper.xss`${eventRecord.name}`}</span>
          `;
            }
        }
    },
    listeners : {
        dataChange : function(event) {
            if (event.store.id === 'events') {
                updateMicrosoftPlanner(event);
            }
        }
    }
});

async function displayUI() {
    await signIn();

    // Hide sign in link and initial UI
    signInLink.style = 'display: none';
    const content = document.getElementById('content');
    content.style = 'display: block';

    // Display calendar after sign in
    const events = await getTasks();
    const calendarEvents = [];
    const resourceID = 1;
    events.value.forEach((event) => {
        const startDateUTC = new Date(event.startDateTime);
        // Convert to local timezone
        const startDateLocal = new Date(
            startDateUTC.getTime() - startDateUTC.getTimezoneOffset() * 60000
        );
        const endDateUTC = new Date(event.dueDateTime);
        // Convert to local timezone
        const endDateLocal = new Date(
            endDateUTC.getTime() - endDateUTC.getTimezoneOffset() * 60000
        );

        calendarEvents.push({
            id              : event.id,
            name            : event.title,
            startDate       : startDateLocal,
            endDate         : endDateLocal,
            taskETag        : event['@odata.etag'].replace(/\\"/g, '"'),
            taskDetailsETag : event?.details['@odata.etag'].replace(/\\"/g, '"'),
            resourceId      : resourceID,
            description     : event.details ? event.details.description : '',
            percentComplete : event?.percentComplete
        });
    });
    calendar.events = calendarEvents;
}

signInLink.addEventListener('click', displayUI);

async function updateMicrosoftPlanner(event) {
    if (event.action == 'update') {
        const id = event.record.id;
        if (`${id}`.startsWith('_generated')) {
            const createTaskRes = await createTask(
                event.record.name,
                event.record.startDate,
                event.record.endDate,
                event.record.description,
                event.record.percentComplete
            );
            // update id and eTags
            calendar.eventStore.applyChangeset({
                updated : [
                    // Will set proper id and eTag for added task
                    {
                        $PhantomId : id,
                        id         : createTaskRes.id,
                        taskETag   : createTaskRes['@odata.etag']
                    }
                ]
            });
            return;
        }

        if (!event.record.taskDetailsETag) return;
        const updateTaskRes = await updateTask(
            id,
            event.record.name,
            event.record.startDate,
            event.record.endDate,
            event.record.taskETag,
            event.record.taskDetailsETag,
            event.record.description,
            event.record.percentComplete
        );
        const updatedObj = {
            id : updateTaskRes.id
        };

        if (updateTaskRes.taskETag) {
            updatedObj.taskETag = updateTaskRes.taskETag;
        }
        if (updateTaskRes.taskDetailsETag) {
            updatedObj.taskDetailsETag = updateTaskRes.taskDetailsETag;
        }

        calendar.eventStore.applyChangeset({
            updated : [
                // Will set proper eTags for updated task
                updatedObj
            ]
        });
    }
    if (event.action == 'remove') {
        const recordsData = event.records.map((record) => record.data);
        recordsData.forEach((record) => {
            if (record.id.startsWith('_generated')) return;
            deleteTask(record.id, record.taskETag);
        });
    }
}

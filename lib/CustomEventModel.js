import { EventModel } from '@bryntum/calendar';

// Custom event model
export default class CustomEventModel extends EventModel {
    static $name = 'CustomEventModel';
    static fields = [
        { name : 'taskETag', type : 'string' },
        { name : 'taskDetailsETag', type : 'string' },
        { name : 'description', type : 'string' },
        { name : 'percentComplete', type : 'number', values : [0, 50, 100] }
    ];
}

import { Client } from '@microsoft/microsoft-graph-client';
import { ensureScope, getToken } from './auth.js';

const authProvider = {
    getAccessToken : async() => {
        return await getToken();
    }
};

// Initialize the Graph client
const graphClient = Client.initWithMiddleware({ authProvider });

export async function getTasks() {
    ensureScope('Tasks.Read');
    return await graphClient
        .api(
      `/planner/plans/${
        import.meta.env.VITE_MICROSOFT_PLANNER_PLAN_ID
      }/tasks?$expand=details`
        )
        .select('id, title, startDateTime, dueDateTime, details, percentComplete')
        .get();
}

export async function createTask(
    name,
    startDate,
    endDate,
    description,
    percentComplete
) {
    ensureScope('Tasks.ReadWrite');
    const task = {
        planId          : import.meta.env.VITE_MICROSOFT_PLANNER_PLAN_ID,
        title           : `${name}`,
        startDateTime   : `${startDate.toISOString()}`,
        dueDateTime     : `${endDate.toISOString()}`,
        details         : { description : description },
        percentComplete : percentComplete
    };
    return await graphClient.api('/planner/tasks').post(task);
}

export async function updateTask(
    id,
    name,
    startDate,
    endDate,
    taskETag,
    taskDetailsETag,
    description,
    percentComplete
) {
    ensureScope('Tasks.ReadWrite');

    const resData = {
        id              : '',
        taskETag        : '',
        taskDetailsETag : ''
    };
    // update task and task details
    const task = {};
    const taskDetails = {};

    if (name) task.title = `${name}`;
    if (startDate) task.startDateTime = `${startDate.toISOString()}`;
    if (endDate) task.dueDateTime = `${endDate.toISOString()}`;
    if (percentComplete) task.percentComplete = percentComplete;

    if (description) taskDetails.description = description;

    // update task only
    if (
        Object.keys(task).length !== 0 &&
      Object.keys(taskDetails).length === 0
    ) {
        const updateTaskRes = await graphClient
            .api(`/planner/tasks/${id}/`)
            .header('If-Match', taskETag)
            .header('prefer', 'return=representation')
            .update(task);

        resData.id = updateTaskRes.id;
        resData.taskETag = updateTaskRes['@odata.etag'];
        return resData;
    }

    // update task details only
    if (
        Object.keys(taskDetails).length !== 0 &&
      Object.keys(task).length === 0
    ) {
        const updateTaskDetailsRes = await graphClient
            .api(`/planner/tasks/${id}/details`)
            .header('If-Match', taskDetailsETag)
            .header('prefer', 'return=representation')
            .update(taskDetails);

        resData.id = updateTaskDetailsRes.id;
        resData.taskDetailsETag = updateTaskDetailsRes['@odata.etag'];
        return resData;
    }

    if (
        Object.keys(task).length !== 0 &&
      Object.keys(taskDetails).length !== 0
    ) {
        // update task and task details
        const updateTaskPromise = graphClient
            .api(`/planner/tasks/${id}/`)
            .header('If-Match', taskETag)
            .header('prefer', 'return=representation')
            .update(task);

        const updateTaskDetailsPromise = graphClient
            .api(`/planner/tasks/${id}/details`)
            .header('If-Match', taskDetailsETag)
            .header('prefer', 'return=representation')
            .update(taskDetails);

        const [updateTaskRes, updateTaskDetailsRes] = await Promise.all([
            updateTaskPromise,
            updateTaskDetailsPromise
        ]);

        resData.id = updateTaskRes.id;
        resData.taskETag = updateTaskRes['@odata.etag'];
        resData.taskDetailsETag = updateTaskDetailsRes['@odata.etag'];

        return resData;
    }
}

export async function deleteTask(id, taskEtag) {
    ensureScope('Tasks.ReadWrite');
    return await graphClient
        .api(`/planner/tasks/${id}`)
        .header('If-Match', taskEtag)
        .delete();
}
// Main function to run the script
function main() {
  const courseId = 123456; // Example course ID
  const discussionTitles = ['Session 2', 'Session 3', 'Session 4', 'Session 6', 'Session 7', 'Session 8'];

  // Get discussions
  getDiscussions(courseId).then(discussions => {
    // Log the discussions
    Logger.log("Discussions: " + JSON.stringify(discussions));

    // Process discussions
    discussions.forEach(discussion => {
      if (discussionTitles.includes(discussion.topicTitle)) {
        Logger.log("Writing to sheet: " + discussion.topicTitle);
        writeToSheet(discussion.topicTitle, discussion.replies);
      }
    });

    // Update last check time
    const now = new Date();
    discussionTitles.forEach(title => {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(title);
      if (sheet) {
        sheet.getRange('G1').setValue(now);
      }
    });
  }).catch(error => {
    Logger.log("Error in getDiscussions: " + error);
  });
}

// Function to get discussion topics IDs
function getDiscussionTopicIds(courseId) {
  const url = `https://courseworks2.columbia.edu/api/v1/courses/${courseId}/discussion_topics`;
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: {
      'Authorization': 'Bearer ' + ScriptProperties.getProperty('CANVAS_API_TOKEN')
    }
  });

  const discussions = JSON.parse(response.getContentText());
  Logger.log("Discussion IDs: " + JSON.stringify(discussions.map(x => x.id)));
  return discussions.map(x => x.id);
}

// Function to get discussions and nested replies
function getDiscussions(courseId) {
  const discussionTopicIds = getDiscussionTopicIds(courseId);

  return Promise.all(discussionTopicIds.map(topicId => {
    return Promise.all([
      getFullDiscussion(courseId, topicId),
      getDiscussionTopic(courseId, topicId)
    ]).then(async ([discussion, topic]) => {
      const topicTitle = topic.title;
      const topicMessage = topic.message;
      const author = topic.author;
      const timestamp = topic.created_at;
      const topicId = topic.id;
      const participants = discussion.participants;
      const replies = discussion.view.length > 0
        ? discussion.view.filter(x => !x.deleted).map(reply => getNestedReplies(reply, participants, topicId))
        : [];
      const flattenedReplies = flatten(replies);
      const repliesWithParents = await getRepliesWithParents(courseId, topicId, flattenedReplies);
      return {
        topicTitle,
        topicMessage,
        id: topicId,
        authorId: author.id || '',
        authorName: author.display_name || '',
        timestamp,
        replies: repliesWithParents
      };
    });
  }));
}

// Function to get full discussion
function getFullDiscussion(courseId, topicId) {
  const url = `https://courseworks2.columbia.edu/api/v1/courses/${courseId}/discussion_topics/${topicId}/view`;
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: {
      'Authorization': 'Bearer ' + ScriptProperties.getProperty('CANVAS_API_TOKEN')
    }
  });

  return JSON.parse(response.getContentText());
}

// Function to get discussion topic
function getDiscussionTopic(courseId, topicId) {
  const url = `https://courseworks2.columbia.edu/api/v1/courses/${courseId}/discussion_topics/${topicId}`;
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: {
      'Authorization': 'Bearer ' + ScriptProperties.getProperty('CANVAS_API_TOKEN')
    }
  });

  return JSON.parse(response.getContentText());
}

// Function to get nested replies
function getNestedReplies(replyObj, participants, topicId) {
  const replies = replyObj.hasOwnProperty('replies')
    ? flatten(replyObj.replies.map(replyObj => getNestedReplies(replyObj, participants, topicId)))
    : [];
  const authorName = participants.find(x => x.id === replyObj.user_id)
    ? participants.find(x => x.id === replyObj.user_id).display_name
    : '';

  return [{
    authorName: authorName,
    message: replyObj.message,
    timestamp: replyObj.created_at,
    parentId: replyObj.parent_id || topicId
  }, ...replies];
}

// Function to get replies with parent details
async function getRepliesWithParents(courseId, topicId, replies) {
  const parentIds = replies.map(reply => reply.parentId).filter(id => id !== topicId);
  if (parentIds.length === 0) return replies;

  const uniqueParentIds = [...new Set(parentIds)];
  const parents = await getEntriesByIds(courseId, topicId, uniqueParentIds);

  return Promise.all(replies.map(async reply => {
    if (reply.parentId !== topicId) {
      const parent = parents.find(parent => parent.id === reply.parentId);
      if (parent) {
        reply.parentAuthor = await getUserName(parent.user_id);
        reply.parentMessage = parent.message;
      }
    } else {
      reply.parentAuthor = 'N/A';
      reply.parentMessage = 'N/A';
    }
    return reply;
  }));
}

// Function to get entries by IDs
function getEntriesByIds(courseId, topicId, ids) {
  const url = `https://courseworks2.columbia.edu/api/v1/courses/${courseId}/discussion_topics/${topicId}/entry_list?${ids.map(id => `ids[]=${id}`).join('&')}`;
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: {
      'Authorization': 'Bearer ' + ScriptProperties.getProperty('CANVAS_API_TOKEN')
    }
  });

  return JSON.parse(response.getContentText());
}

// Function to get user name by user ID
function getUserName(userId) {
  const url = `https://courseworks2.columbia.edu/api/v1/users/${userId}/profile`;
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: {
      'Authorization': 'Bearer ' + ScriptProperties.getProperty('CANVAS_API_TOKEN')
    }
  });

  const userProfile = JSON.parse(response.getContentText());
  return userProfile.name;
}

// Utility function to flatten arrays
function flatten(arr) {
  return arr.reduce((acc, cur) => Array.isArray(cur) ? [...acc, ...cur] : [...acc, cur], []);
}

// Function to write data to Google Sheets
function writeToSheet(sheetName, replies) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    Logger.log("Sheet not found: " + sheetName);
    return;
  }

  // Check if the header is already present
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Name', 'Comment', 'Submission Time', 'Reply To', 'Comment Replied To']);
  }

  // Get the last check time from cell G1
  let lastCheckTimeCell = sheet.getRange('G1').getValue();
  if (lastCheckTimeCell) {
    lastCheckTime = new Date(lastCheckTimeCell);
  } else {
    lastCheckTime = new Date(0); // Default to the epoch if G1 is empty
  }
  
  replies.forEach(reply => {
    Logger.log("Reply object: " + JSON.stringify(reply));
    const replyDate = new Date(reply.timestamp);
    Logger.log("Reply timestamp: " + reply.timestamp);
    Logger.log("Last check time: " + lastCheckTime.toISOString());
    if (replyDate > lastCheckTime && reply.authorName && reply.message) {
      const replyTo = reply.parentId === reply.id ? 'N/A' : reply.parentAuthor;
      const commentRepliedTo = reply.parentId === reply.id ? 'N/A' : stripHTML(reply.parentMessage);
      const formattedDate = Utilities.formatDate(replyDate, Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss");
      Logger.log("Appending row: " + [reply.authorName, stripHTML(reply.message), formattedDate, replyTo, commentRepliedTo].join(", "));
      sheet.appendRow([reply.authorName, stripHTML(reply.message), formattedDate, replyTo, commentRepliedTo]);
    }
  });
}

// Function to create a time-driven trigger
function createTrigger() {
  ScriptApp.newTrigger('main')
    .timeBased()
    .everyMinutes(5)
    .create();
}

// Function to strip HTML from a string
function stripHTML(html) {
  return html ? html.replace(/(<([^>]+)>)/gi, "").replace(/&nbsp;/g, " ") : '';
}
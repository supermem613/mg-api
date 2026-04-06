#!/usr/bin/env node
'use strict';

const { graphFetch } = require('./mg-fetch');

const GRAPH = 'https://graph.microsoft.com/v1.0';
const OUTLOOK = 'https://outlook.office.com/api/v2.0';

function headers(token, contentType) {
  const h = { Authorization: `Bearer ${token}`, Accept: 'application/json' };
  if (contentType) h['Content-Type'] = contentType;
  return h;
}

function qs(params) {
  const parts = [];
  for (const [k, v] of Object.entries(params)) {
    if (v !== undefined && v !== null) parts.push(`$${k}=${encodeURIComponent(v)}`);
  }
  return parts.length ? '?' + parts.join('&') : '';
}

async function get(token, path, queryParams = {}) {
  const url = `${GRAPH}${path}${qs(queryParams)}`;
  return graphFetch(url, { headers: headers(token) });
}

async function post(token, path, body, base = GRAPH) {
  const url = `${base}${path}`;
  return graphFetch(url, {
    method: 'POST',
    headers: headers(token, 'application/json'),
    body: JSON.stringify(body),
  });
}

async function patch(token, path, body) {
  const url = `${GRAPH}${path}`;
  return graphFetch(url, {
    method: 'PATCH',
    headers: headers(token, 'application/json'),
    body: JSON.stringify(body),
  });
}

async function del(token, path) {
  const url = `${GRAPH}${path}`;
  return graphFetch(url, { method: 'DELETE', headers: headers(token) });
}

// ---------------------------------------------------------------------------
// Email
// ---------------------------------------------------------------------------

async function listMessages(token, { top, filter, select, orderby } = {}) {
  return get(token, '/me/messages', { top, filter, select, orderby });
}

async function getMessage(token, { messageId, select } = {}) {
  return get(token, `/me/messages/${messageId}`, { select });
}

async function sendEmail(outlookToken, { to, cc, bcc, subject, body, bodyType = 'Text', saveToSentItems = true }) {
  const message = {
    Message: {
      Subject: subject,
      Body: { ContentType: bodyType, Content: body },
      ToRecipients: to.map(e => ({ EmailAddress: { Address: e } })),
      ...(cc && { CcRecipients: cc.map(e => ({ EmailAddress: { Address: e } })) }),
      ...(bcc && { BccRecipients: bcc.map(e => ({ EmailAddress: { Address: e } })) }),
    },
    SaveToSentItems: saveToSentItems,
  };
  return post(outlookToken, '/me/sendmail', message, OUTLOOK);
}

async function replyToMessage(outlookToken, { messageId, comment }) {
  return post(outlookToken, `/me/messages/${messageId}/reply`, { Comment: comment }, OUTLOOK);
}

async function searchMessages(token, { query, top } = {}) {
  const params = {};
  if (top) params.top = top;
  const search = encodeURIComponent(`"${query}"`);
  const base = `${GRAPH}/me/messages?$search=${search}`;
  const suffix = Object.keys(params).length ? '&' + Object.entries(params).map(([k, v]) => `$${k}=${v}`).join('&') : '';
  return graphFetch(base + suffix, { headers: headers(token) });
}

async function moveMessage(token, { messageId, destinationFolderId }) {
  return post(token, `/me/messages/${messageId}/move`, { destinationId: destinationFolderId });
}

async function deleteMessage(token, { messageId }) {
  return del(token, `/me/messages/${messageId}`);
}

async function listAttachments(token, { messageId }) {
  return get(token, `/me/messages/${messageId}/attachments`);
}

// ---------------------------------------------------------------------------
// Calendar
// ---------------------------------------------------------------------------

async function listEvents(token, { top, filter, select, orderby } = {}) {
  return get(token, '/me/events', { top, filter, select, orderby });
}

async function getEvent(token, { eventId, select } = {}) {
  return get(token, `/me/events/${eventId}`, { select });
}

async function createEvent(token, { subject, body, start, end, attendees, location, isOnlineMeeting }) {
  const event = {
    subject,
    body: body ? { contentType: 'Text', content: body } : undefined,
    start: typeof start === 'string' ? { dateTime: start, timeZone: 'UTC' } : start,
    end: typeof end === 'string' ? { dateTime: end, timeZone: 'UTC' } : end,
    ...(attendees && {
      attendees: attendees.map(a =>
        typeof a === 'string'
          ? { emailAddress: { address: a }, type: 'required' }
          : a
      ),
    }),
    ...(location && { location: typeof location === 'string' ? { displayName: location } : location }),
    ...(isOnlineMeeting !== undefined && { isOnlineMeeting }),
  };
  return post(token, '/me/events', event);
}

async function updateEvent(token, { eventId, ...fields }) {
  return patch(token, `/me/events/${eventId}`, fields);
}

async function deleteEvent(token, { eventId }) {
  return del(token, `/me/events/${eventId}`);
}

async function acceptEvent(token, { eventId, comment, sendResponse = true }) {
  return post(token, `/me/events/${eventId}/accept`, { comment, sendResponse });
}

async function declineEvent(token, { eventId, comment, sendResponse = true }) {
  return post(token, `/me/events/${eventId}/decline`, { comment, sendResponse });
}

async function findMeetingTimes(token, { attendees, timeConstraint, meetingDuration }) {
  const payload = {
    ...(attendees && {
      attendees: attendees.map(a =>
        typeof a === 'string'
          ? { emailAddress: { address: a }, type: 'required' }
          : a
      ),
    }),
    ...(timeConstraint && { timeConstraint }),
    ...(meetingDuration && { meetingDuration }),
  };
  return post(token, '/me/findMeetingTimes', payload);
}

// ---------------------------------------------------------------------------
// Teams
// ---------------------------------------------------------------------------

async function listJoinedTeams(token) {
  return get(token, '/me/joinedTeams');
}

async function listChannels(token, { teamId }) {
  return get(token, `/teams/${teamId}/channels`);
}

async function sendChannelMessage(token, { teamId, channelId, content, contentType = 'text' }) {
  return post(token, `/teams/${teamId}/channels/${channelId}/messages`, {
    body: { contentType, content },
  });
}

async function listChats(outlookToken, { top } = {}) {
  const params = top ? `?$top=${top}` : '';
  return graphFetch(`${OUTLOOK}/me/chats${params}`, { headers: headers(outlookToken) });
}

async function getChatMessages(outlookToken, { chatId, top } = {}) {
  const params = top ? `?$top=${top}` : '';
  return graphFetch(`${OUTLOOK}/me/chats/${chatId}/messages${params}`, { headers: headers(outlookToken) });
}

async function sendChatMessage(token, { chatId, content, contentType = 'text' }) {
  return post(token, `/chats/${chatId}/messages`, {
    body: { contentType, content },
  });
}

// ---------------------------------------------------------------------------
// Users
// ---------------------------------------------------------------------------

async function getMyProfile(token, { select } = {}) {
  return get(token, '/me', { select });
}

async function searchPeople(token, { query, top } = {}) {
  const search = encodeURIComponent(`"${query}"`);
  const params = top ? `&$top=${top}` : '';
  return graphFetch(`${GRAPH}/me/people?$search=${search}${params}`, { headers: headers(token) });
}

async function getUser(token, { userId, select } = {}) {
  return get(token, `/users/${userId}`, { select });
}

// ---------------------------------------------------------------------------
// Exports
// ---------------------------------------------------------------------------

module.exports = {
  // Email
  listMessages,
  getMessage,
  sendEmail,
  replyToMessage,
  searchMessages,
  moveMessage,
  deleteMessage,
  listAttachments,
  // Calendar
  listEvents,
  getEvent,
  createEvent,
  updateEvent,
  deleteEvent,
  acceptEvent,
  declineEvent,
  findMeetingTimes,
  // Teams
  listJoinedTeams,
  listChannels,
  sendChannelMessage,
  listChats,
  getChatMessages,
  sendChatMessage,
  // Users
  getMyProfile,
  searchPeople,
  getUser,
};

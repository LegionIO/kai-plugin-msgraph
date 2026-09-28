/**
 * Teams system-event messages ("X removed Y", "Meeting started", "Recording
 * available", …) grouped into user-configurable categories. A category that is
 * not tracked never marks a chat unread or fires `message-received`.
 *
 * The same event arrives in two shapes: Graph chat previews carry an
 * `eventDetail['@odata.type']`, while Trouter pushes carry the IC3 `messagetype`
 * (plus XML content for calls). Both are mapped here.
 */

export type SystemEventCategory =
  | 'membersAdded'
  | 'membersRemoved'
  | 'callStarted'
  | 'callEnded'
  | 'callRecording'
  | 'callTranscript'
  | 'chatRenamed'
  | 'messagePinned'
  | 'other';

export type TrackedSystemEvents = Record<SystemEventCategory, boolean>;

export const SYSTEM_EVENT_LABELS: Record<SystemEventCategory, string> = {
  membersAdded: 'Members added / joined',
  membersRemoved: 'Members removed / left',
  callStarted: 'Meeting / call started',
  callEnded: 'Meeting / call ended',
  callRecording: 'Recording available',
  callTranscript: 'Transcript available',
  chatRenamed: 'Chat renamed',
  messagePinned: 'Message pinned',
  other: 'Other system events (meeting options, tabs, etc.)',
};

/** Default: system events are silent — only real messages mark a chat unread. */
export const DEFAULT_TRACKED_SYSTEM_EVENTS: TrackedSystemEvents = {
  membersAdded: false,
  membersRemoved: false,
  callStarted: false,
  callEnded: false,
  callRecording: false,
  callTranscript: false,
  chatRenamed: false,
  messagePinned: false,
  other: false,
};

export function resolveTrackedSystemEvents(partial?: Partial<TrackedSystemEvents> | null): TrackedSystemEvents {
  return { ...DEFAULT_TRACKED_SYSTEM_EVENTS, ...(partial ?? {}) };
}

/** Category for a Graph `eventDetail` (`#microsoft.graph.fooEventMessageDetail`). */
export function graphEventCategory(odataType: string | null | undefined): SystemEventCategory {
  const t = String(odataType ?? '').replace('#microsoft.graph.', '');
  switch (t) {
    case 'membersAddedEventMessageDetail':
    case 'membersJoinedEventMessageDetail':
      return 'membersAdded';
    case 'membersDeletedEventMessageDetail':
    case 'membersLeftEventMessageDetail':
      return 'membersRemoved';
    case 'callStartedEventMessageDetail':
      return 'callStarted';
    case 'callEndedEventMessageDetail':
      return 'callEnded';
    case 'callRecordingEventMessageDetail':
      return 'callRecording';
    case 'callTranscriptEventMessageDetail':
      return 'callTranscript';
    case 'chatRenamedEventMessageDetail':
      return 'chatRenamed';
    case 'messagePinnedEventMessageDetail':
    case 'messageUnpinnedEventMessageDetail':
      return 'messagePinned';
    default:
      return 'other';
  }
}

/**
 * Category for an IC3 push message, or null when it is a regular chat message
 * (Text, RichText/Html, RichText/Media_Card, …).
 */
export function ic3EventCategory(messagetype: string, content: string | null | undefined): SystemEventCategory | null {
  const mt = messagetype;
  if (mt === 'RichText/Media_CallRecording') return 'callRecording';
  if (mt === 'RichText/Media_CallTranscript') return 'callTranscript';
  if (mt === 'Event/Call') {
    return /<partlist\b[^>]*\btype="ended"/i.test(content ?? '') ? 'callEnded' : 'callStarted';
  }
  if (mt.startsWith('ThreadActivity/')) {
    const kind = mt.slice('ThreadActivity/'.length);
    if (kind === 'AddMember' || kind === 'MemberJoined') return 'membersAdded';
    if (kind === 'DeleteMember' || kind === 'MemberLeft') return 'membersRemoved';
    if (kind === 'TopicUpdate') return 'chatRenamed';
    if (/pin/i.test(kind)) return 'messagePinned';
    return 'other';
  }
  if (mt.startsWith('Event/')) return 'other';
  return null;
}

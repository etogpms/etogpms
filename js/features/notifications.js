// Smart Notification System for Calendar Activities
// Generates 24-hour-prior reminders for meetings, activities, and leaves.
// Notifications appear in the bell icon dropdown.
(function (window) {
    if (window.NotificationsFeature) return;

    let notificationService;
    let calendarService;
    let auth;
    let notifications = [];
    let unsubscribeNotifications = null;
    let notificationCheckInterval = null;
    let panelOpen = false;

    // DOM references
    let bellBtn = null;
    let badgeEl = null;
    let panelEl = null;
    let listEl = null;
    let emptyEl = null;

    const NotificationsFeature = {
        init({ notificationService: ns, calendarService: cs, auth: authInstance }) {
            notificationService = ns;
            calendarService = cs;
            auth = authInstance;

            bellBtn = document.getElementById('headerNotificationsBtn');
            badgeEl = document.getElementById('notifBadgeCount');
            panelEl = document.getElementById('notificationsPanel');
            listEl = document.getElementById('notificationsList');
            emptyEl = document.getElementById('notificationsEmpty');

            if (bellBtn) {
                bellBtn.addEventListener('click', (e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    togglePanel();
                });
            }

            // Close panel when clicking outside
            document.addEventListener('click', (e) => {
                if (panelOpen && panelEl && !panelEl.contains(e.target) && !bellBtn.contains(e.target)) {
                    closePanel();
                }
            });

            // Mark all read button
            const markAllBtn = document.getElementById('notifMarkAllRead');
            if (markAllBtn) {
                markAllBtn.addEventListener('click', markAllAsRead);
            }

            return this;
        },

        /**
         * Called when user signs in - start listening for notifications
         */
        startListening(user) {
            if (!user) return;
            stopListening();

            // Subscribe to notifications for this user, ordered by creation time
            unsubscribeNotifications = notificationService.collection()
                .where('recipientUid', '==', user.uid)
                .orderBy('createdAt', 'desc')
                .limit(50)
                .onSnapshot((snap) => {
                    notifications = snap.docs.map(doc => ({ id: doc.id, ...doc.data() }));
                    renderNotifications();
                    updateBadge();
                }, (err) => {
                    console.error('[Notifications] subscribe error:', err);
                });

            // Periodically check for events needing notifications (every 15 minutes)
            generateUpcomingNotifications(user);
            notificationCheckInterval = setInterval(() => {
                generateUpcomingNotifications(user);
            }, 15 * 60 * 1000); // 15 minutes
        },

        /**
         * Called when user signs out - stop listening
         */
        stopListening() {
            stopListening();
        }
    };

    function stopListening() {
        if (unsubscribeNotifications) {
            unsubscribeNotifications();
            unsubscribeNotifications = null;
        }
        if (notificationCheckInterval) {
            clearInterval(notificationCheckInterval);
            notificationCheckInterval = null;
        }
        notifications = [];
        renderNotifications();
        updateBadge();
    }

    /**
     * Check all calendar_activities and generate notification docs for events
     * happening within the next 24 hours that haven't been notified yet.
     */
    async function generateUpcomingNotifications(user) {
        if (!user) return;

        // OPTIMIZATION: Throttle checks to once every 10 mins per session (even on refresh)
        const now = new Date();
        const lastCheck = localStorage.getItem(`notif_last_check_${user.uid}`);
        if (lastCheck && (now.getTime() - parseInt(lastCheck, 10) < 10 * 60 * 1000)) {
            return; // Checked recently, skip
        }
        localStorage.setItem(`notif_last_check_${user.uid}`, now.getTime().toString());

        try {
            const in24h = new Date(now.getTime() + 24 * 60 * 60 * 1000);

            // Get the date strings for querying
            const todayStr = formatDateStr(now);
            const tomorrowStr = formatDateStr(in24h);

            // OPTIMIZATION: Filter by ownerUid in the query to reduce read counts
            const snapshot = await calendarService.collection()
                .where('ownerUid', '==', user.uid)
                .where('startDate', '>=', todayStr)
                .where('startDate', '<=', tomorrowStr)
                .get();

            if (snapshot.empty) return;

            const events = snapshot.docs.map(d => ({ id: d.id, ...d.data() }));

            for (const ev of events) {
                // Build the event's actual start datetime
                const eventStart = buildEventDateTime(ev.startDate, ev.startTime, ev.allDay);
                if (!eventStart) continue;

                const timeDiff = eventStart.getTime() - now.getTime();

                // Only create notification if event is within the next 24 hours and hasn't passed
                if (timeDiff <= 0 || timeDiff > 24 * 60 * 60 * 1000) continue;

                // Check if a notification already exists for this event+user
                const notifId = `notif_${ev.id}_${user.uid}`;

                try {
                    const existingDoc = await notificationService.collection().doc(notifId).get();
                    if (existingDoc.exists) continue; // Already notified
                } catch (_) {
                    // ignore, proceed to create
                }

                // Determine who should be notified:
                // The event owner gets the notification
                if (ev.ownerUid !== user.uid) continue;

                // Build notification message
                const typeLabel = (ev.type || 'activity').charAt(0).toUpperCase() + (ev.type || 'activity').slice(1);
                const timeStr = ev.allDay ? 'All Day' : formatTime12(ev.startTime);
                const dateFormatted = formatDisplayDate(ev.startDate);

                const hoursLeft = Math.round(timeDiff / (1000 * 60 * 60));
                const minutesLeft = Math.round(timeDiff / (1000 * 60));
                let timeAway = '';
                if (hoursLeft > 0) {
                    timeAway = `in ${hoursLeft} hour${hoursLeft === 1 ? '' : 's'}`;
                } else {
                    timeAway = `in ${minutesLeft} minute${minutesLeft === 1 ? '' : 's'}`;
                }

                const notifData = {
                    id: notifId,
                    recipientUid: user.uid,
                    recipientEmail: user.email || '',
                    eventId: ev.id,
                    eventTitle: ev.title || 'Untitled Event',
                    eventType: ev.type || 'activity',
                    eventDate: ev.startDate,
                    eventTime: ev.startTime || null,
                    eventAllDay: !!ev.allDay,
                    message: `📅 Reminder: "${ev.title}" (${typeLabel}) is coming up ${timeAway} — ${dateFormatted} at ${timeStr}`,
                    shortMessage: `${typeLabel}: "${ev.title}" — ${dateFormatted}, ${timeStr}`,
                    read: false,
                    createdAt: window.AppFirebase.serverTimestamp(),
                };

                try {
                    await notificationService.collection().doc(notifId).set(notifData);
                    console.log('[Notifications] Created reminder for:', ev.title);
                } catch (err) {
                    console.error('[Notifications] Failed to create notification:', err);
                }
            }
        } catch (err) {
            console.error('[Notifications] Error checking upcoming events:', err);
        }
    }

    function buildEventDateTime(dateStr, timeStr, allDay) {
        if (!dateStr) return null;
        const [y, m, d] = dateStr.split('-').map(Number);
        if (allDay || !timeStr) {
            // For all-day events, consider the start as 8:00 AM of that day
            return new Date(y, m - 1, d, 8, 0, 0);
        }
        const [h, min] = timeStr.split(':').map(Number);
        return new Date(y, m - 1, d, h, min, 0);
    }

    function formatDateStr(date) {
        const y = date.getFullYear();
        const m = String(date.getMonth() + 1).padStart(2, '0');
        const d = String(date.getDate()).padStart(2, '0');
        return `${y}-${m}-${d}`;
    }

    function formatDisplayDate(dateStr) {
        if (!dateStr) return '';
        const [y, m, d] = dateStr.split('-').map(Number);
        const date = new Date(y, m - 1, d);
        return date.toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric', year: 'numeric' });
    }

    function formatTime12(timeStr) {
        if (!timeStr) return '';
        const [h, m] = timeStr.split(':');
        const hour = parseInt(h, 10);
        const ampm = hour >= 12 ? 'PM' : 'AM';
        const h12 = hour % 12 || 12;
        return `${h12}:${m} ${ampm}`;
    }

    function timeAgoStr(timestamp) {
        if (!timestamp) return '';
        let date;
        if (timestamp.toDate) {
            date = timestamp.toDate();
        } else if (timestamp instanceof Date) {
            date = timestamp;
        } else {
            return '';
        }
        const now = new Date();
        const diff = now.getTime() - date.getTime();
        const mins = Math.floor(diff / 60000);
        if (mins < 1) return 'Just now';
        if (mins < 60) return `${mins}m ago`;
        const hours = Math.floor(mins / 60);
        if (hours < 24) return `${hours}h ago`;
        const days = Math.floor(hours / 24);
        if (days === 1) return 'Yesterday';
        if (days < 7) return `${days}d ago`;
        return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
    }

    function togglePanel() {
        if (panelOpen) {
            closePanel();
        } else {
            openPanel();
        }
    }

    function openPanel() {
        if (panelEl) {
            panelEl.style.display = 'block';
            panelOpen = true;
        }
    }

    function closePanel() {
        if (panelEl) {
            panelEl.style.display = 'none';
            panelOpen = false;
        }
    }

    function updateBadge() {
        const unread = notifications.filter(n => !n.read).length;
        if (badgeEl) {
            if (unread > 0) {
                badgeEl.textContent = unread > 99 ? '99+' : unread;
                badgeEl.style.display = 'flex';
            } else {
                badgeEl.style.display = 'none';
            }
        }
    }

    function renderNotifications() {
        if (!listEl) return;

        if (notifications.length === 0) {
            listEl.innerHTML = '';
            if (emptyEl) emptyEl.style.display = 'block';
            return;
        }

        if (emptyEl) emptyEl.style.display = 'none';

        let html = '';
        notifications.forEach(n => {
            const typeIcon = getTypeIcon(n.eventType);
            const typeColor = getTypeColor(n.eventType);
            const readClass = n.read ? 'notif-item-read' : 'notif-item-unread';
            const timeAgo = timeAgoStr(n.createdAt);

            html += `
        <div class="notif-item ${readClass}" data-id="${n.id}" onclick="window.NotificationsFeature.onClickNotification('${n.id}')">
          <div class="notif-icon-wrap" style="background: ${typeColor}15; color: ${typeColor};">
            <i class="${typeIcon}"></i>
          </div>
          <div class="notif-content">
            <div class="notif-message">${escapeHtml(n.shortMessage || n.message || '')}</div>
            <div class="notif-time">${timeAgo}</div>
          </div>
          ${!n.read ? '<div class="notif-unread-dot"></div>' : ''}
        </div>
      `;
        });

        listEl.innerHTML = html;
    }

    function getTypeIcon(type) {
        switch (type) {
            case 'meeting': return 'fa-solid fa-handshake';
            case 'leave': return 'fa-solid fa-plane-departure';
            case 'activity': return 'fa-solid fa-calendar-check';
            default: return 'fa-solid fa-bell';
        }
    }

    function getTypeColor(type) {
        switch (type) {
            case 'meeting': return '#0d6efd';
            case 'leave': return '#ffc107';
            case 'activity': return '#198754';
            default: return '#6c757d';
        }
    }

    function escapeHtml(str) {
        return (str || '').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    }

    async function markAllAsRead() {
        const user = auth.currentUser;
        if (!user) return;

        const unread = notifications.filter(n => !n.read);
        if (unread.length === 0) return;

        try {
            const batch = notificationService.batch();
            unread.forEach(n => {
                const ref = notificationService.collection().doc(n.id);
                batch.update(ref, { read: true });
            });
            await batch.commit();
        } catch (err) {
            console.error('[Notifications] Mark all read error:', err);
        }
    }

    // Exposed globally so onclick works
    window.NotificationsFeature = Object.assign(NotificationsFeature, {
        onClickNotification: async (notifId) => {
            // Mark as read
            const notif = notifications.find(n => n.id === notifId);
            if (notif && !notif.read) {
                try {
                    await notificationService.collection().doc(notifId).update({ read: true });
                } catch (err) {
                    console.error('[Notifications] mark read error:', err);
                }
            }

            // Navigate to calendar tab and try to open the event
            closePanel();
            if (notif && notif.eventId) {
                // Switch to calendar tab if available
                const calTab = document.getElementById('calendarTab');
                if (calTab) calTab.click();

                // Slight delay then try to open the event
                setTimeout(() => {
                    if (window.Calendar && window.Calendar.openEvent) {
                        window.Calendar.openEvent(notif.eventId);
                    }
                }, 500);
            }
        }
    });

})(window);

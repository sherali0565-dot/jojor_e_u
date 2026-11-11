
def should_wait_for_campaign_interval(history_ws, website_rows):
    """
    A new batch of pending tasks should get its first email immediately.
    After that first successful email, enforce the 12-minute interval.
    """

    pending_dates = []

    for row in website_rows:
        if str(row.get("Status", "")).strip().upper() != "PENDING":
            continue

        created_at = str(row.get("Created At", "")).strip()

        if not created_at:
            continue

        try:
            created_time = datetime.strptime(
                created_at,
                "%Y-%m-%d %H:%M:%S"
            )
            pending_dates.append(created_time)
        except ValueError:
            continue

    last_send = last_successful_send(history_ws)

    # No previous successful email: send immediately.
    if last_send is None:
        return False

    # If pending tasks were created after the last successful send,
    # this is a new batch: allow its first email immediately.
    if pending_dates:
        campaign_start = min(pending_dates)

        if last_send < campaign_start:
            return False

    # Otherwise enforce the normal 12-minute interval.
    elapsed = (now_local() - last_send).total_seconds()

    return elapsed < SEND_INTERVAL_SECONDS

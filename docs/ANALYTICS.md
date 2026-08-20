# Product analytics

Pre-Launch Beta 1 includes a privacy-minimized analytics system for product development.

## What is recorded

With user consent, the app can record product events such as:

- app start
- feature used
- update check
- session heartbeat
- session end
- website page view
- website download click
- Word add-in open
- Word add-in action

The desktop app may use a random anonymous installation UUID so repeat usage can be counted without requiring an account.

## What is never recorded by telemetry

- manuscript text
- filename
- file path
- citation text
- reference text
- figure/table content
- manuscript hashes
- DOCX metadata

The telemetry client uses an allow-list for event properties so arbitrary manuscript-derived values are not accepted.

## Feedback

Feedback is a separate table because it is explicit user-submitted content.

The feedback form accepts:

- 1-5 rating
- category
- message
- optional email
- optional contact permission

## Storage

Analytics are stored in Supabase tables:

- `wjmc_events`
- `wjmc_feedback`

Row Level Security allows anonymous/public clients to insert records only. Anonymous/public reads, updates, and deletes are denied.

## Owner reporting

The Supabase project owner can use the SQL editor to summarize usage.

### Sessions and average duration

```sql
select
  count(*) filter (where event_name = 'session_end') as completed_sessions,
  round(avg(duration_seconds) filter (where event_name = 'session_end')) as avg_session_seconds
from public.wjmc_events;
```

### Feature usage

```sql
select
  properties ->> 'feature' as feature,
  count(*) as uses
from public.wjmc_events
where event_name = 'feature_used'
group by 1
order by uses desc;
```

### Download clicks

```sql
select
  properties ->> 'asset' as asset,
  count(*) as clicks
from public.wjmc_events
where event_name = 'download_click'
group by 1
order by clicks desc;
```

GitHub Release asset download counts should be treated as the source of truth for completed GitHub asset downloads. Website `download_click` events measure user intent/click-through.

### Feedback summary

```sql
select
  category,
  count(*) as responses,
  round(avg(rating), 2) as avg_rating
from public.wjmc_feedback
group by category
order by responses desc;
```

## Consent

Desktop analytics are opt-in. The website and Word add-in have separate consent controls. Feedback submission is always explicit.

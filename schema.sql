-- ============================================================
--  מבצעי חנוכה בבתי הרפואה · פנסאים × לב חב"ד
--  מבנה בסיס נתונים ל־Supabase
--  להריץ פעם אחת ב־SQL Editor של פרויקט חדש.
-- ============================================================

create extension if not exists pgcrypto;

-- ---------- בתי רפואה ----------
create table if not exists hospitals (
  id          uuid primary key default gen_random_uuid(),
  name        text not null,
  city        text,
  sort        int  not null default 0,
  active      boolean not null default true,
  created_at  timestamptz not null default now()
);

-- ---------- משמרות (יום + שעות בבית רפואה) ----------
-- כל שורה = משמרת אחת. ביום אפשר להוסיף כמה משמרות, ולכל בית רפואה שעות משלו.
create table if not exists shifts (
  id           uuid primary key default gen_random_uuid(),
  hospital_id  uuid not null references hospitals(id) on delete cascade,
  date         date not null,
  heb_date     text not null,              -- למשל: כ"ו כסלו
  start_time   time not null,
  end_time     time,
  capacity     int  not null default 10,   -- מספר אבות מקסימלי
  active       boolean not null default true,
  created_at   timestamptz not null default now()
);
create index if not exists shifts_hospital_idx on shifts(hospital_id, date);

-- ---------- רשימות סגורות (ערים, כיתות, סניפים) ----------
create table if not exists options (
  id     serial primary key,
  kind   text not null check (kind in ('city','grade','branch')),
  label  text not null,
  sort   int  not null default 0,
  unique (kind, label)
);

-- ---------- הרשמות (אבא) ----------
create table if not exists registrations (
  id             uuid primary key default gen_random_uuid(),
  shift_id       uuid not null references shifts(id),
  father_first   text not null,
  father_last    text not null,
  father_mother  text not null,
  phone          text not null,
  city           text not null,
  status         text not null default 'active' check (status in ('active','cancelled')),
  admin_note     text,
  created_at     timestamptz not null default now(),
  cancelled_at   timestamptz
);
create index if not exists registrations_shift_idx on registrations(shift_id) where status = 'active';

-- ---------- ילדים ----------
create table if not exists children (
  id               uuid primary key default gen_random_uuid(),
  registration_id  uuid not null references registrations(id) on delete cascade,
  first_name       text not null,
  last_name        text not null,
  mother_name      text not null,
  grade            text not null,
  branch           text not null,
  sort             int  not null default 0
);
create index if not exists children_reg_idx on children(registration_id);

-- ---------- מנהלים ----------
-- רק מייל שמופיע כאן יכול לראות ולנהל הרשמות.
create table if not exists admins (
  email text primary key
);

create or replace function is_admin() returns boolean
language sql stable security definer set search_path = public as $$
  select exists (select 1 from admins where lower(email) = lower(auth.jwt() ->> 'email'));
$$;

-- ============================================================
--  זמינות: מחזירה רק "מלא/פנוי" לכל משמרת (בלי לחשוף נרשמים)
-- ============================================================
create or replace function get_availability()
returns table (shift_id uuid, is_full boolean)
language sql stable security definer set search_path = public as $$
  select s.id,
         (select count(*) from registrations r
           where r.shift_id = s.id and r.status = 'active') >= s.capacity
  from shifts s
  where s.active;
$$;

-- ============================================================
--  הרשמה: נועלת את המשמרת, בודקת מכסה, ורק אז שומרת.
--  שני אבות שלוחצים באותה שנייה על המקום העשירי — רק אחד ייכנס.
-- ============================================================
create or replace function register(p_shift uuid, p_father jsonb, p_children jsonb)
returns uuid
language plpgsql security definer set search_path = public as $$
declare
  v_cap   int;
  v_taken int;
  v_id    uuid;
  v_child jsonb;
  v_i     int := 0;
begin
  select capacity into v_cap from shifts
   where id = p_shift and active
   for update;                                   -- נעילה
  if not found then raise exception 'SHIFT_CLOSED'; end if;

  select count(*) into v_taken from registrations
   where shift_id = p_shift and status = 'active';
  if v_taken >= v_cap then raise exception 'SHIFT_FULL'; end if;

  if coalesce(trim(p_father->>'first'),'')  = '' or
     coalesce(trim(p_father->>'last'),'')   = '' or
     coalesce(trim(p_father->>'mother'),'') = '' or
     coalesce(trim(p_father->>'phone'),'')  = '' or
     coalesce(trim(p_father->>'city'),'')   = '' then
    raise exception 'MISSING_FIELDS';
  end if;
  if jsonb_array_length(coalesce(p_children,'[]'::jsonb)) = 0 then
    raise exception 'NO_CHILDREN';
  end if;

  insert into registrations (shift_id, father_first, father_last, father_mother, phone, city)
  values (p_shift,
          trim(p_father->>'first'), trim(p_father->>'last'), trim(p_father->>'mother'),
          trim(p_father->>'phone'), trim(p_father->>'city'))
  returning id into v_id;

  for v_child in select * from jsonb_array_elements(p_children) loop
    if coalesce(trim(v_child->>'first'),'') = '' or coalesce(trim(v_child->>'grade'),'') = ''
       or coalesce(trim(v_child->>'branch'),'') = '' then
      raise exception 'MISSING_FIELDS';
    end if;
    insert into children (registration_id, first_name, last_name, mother_name, grade, branch, sort)
    values (v_id, trim(v_child->>'first'), trim(v_child->>'last'), trim(v_child->>'mother'),
            trim(v_child->>'grade'), trim(v_child->>'branch'), v_i);
    v_i := v_i + 1;
  end loop;

  return v_id;
end;
$$;

-- ============================================================
--  הרשאות (RLS)
-- ============================================================
alter table hospitals     enable row level security;
alter table shifts        enable row level security;
alter table options       enable row level security;
alter table registrations enable row level security;
alter table children      enable row level security;
alter table admins        enable row level security;

-- ציבור: קריאה בלבד של בתי רפואה, משמרות ורשימות
create policy pub_read_hospitals on hospitals for select using (active or is_admin());
create policy pub_read_shifts    on shifts    for select using (active or is_admin());
create policy pub_read_options   on options   for select using (true);

-- מנהלים: הכל
create policy adm_hospitals on hospitals     for all using (is_admin()) with check (is_admin());
create policy adm_shifts    on shifts        for all using (is_admin()) with check (is_admin());
create policy adm_options   on options       for all using (is_admin()) with check (is_admin());
create policy adm_regs      on registrations for all using (is_admin()) with check (is_admin());
create policy adm_children  on children      for all using (is_admin()) with check (is_admin());
create policy adm_admins    on admins        for select using (is_admin());

-- ציבור נרשם רק דרך הפונקציה, לא ישירות לטבלאות
revoke all on registrations, children, admins from anon;
grant execute on function register(uuid, jsonb, jsonb) to anon, authenticated;
grant execute on function get_availability()           to anon, authenticated;

-- ============================================================
--  נתוני התחלה
-- ============================================================
insert into hospitals (name, city, sort) values
  ('בית החולים מאיר', 'כפר סבא', 1),
  ('בית החולים וולפסון', 'חולון', 2),
  ('בית החולים סורוקה', 'באר שבע', 3),
  ('בית החולים קפלן', 'רחובות', 4);

-- כ"ו–ל' כסלו תשפ"ז = 6–10 בדצמבר 2026, 16:30–18:30, 10 אבות
insert into shifts (hospital_id, date, heb_date, start_time, end_time, capacity)
select h.id, d.date, d.heb, '16:30', '18:30', 10
from hospitals h
cross join (values
  ('2026-12-06'::date, 'כ"ו כסלו'),
  ('2026-12-07'::date, 'כ"ז כסלו'),
  ('2026-12-08'::date, 'כ"ח כסלו'),
  ('2026-12-09'::date, 'כ"ט כסלו'),
  ('2026-12-10'::date, 'ל'' כסלו')
) as d(date, heb);

-- רשימות זמניות (להחליף בנתונים אמיתיים)
insert into options (kind, label, sort) values
  ('city','בני ברק',1),('city','ירושלים',2),('city','כפר חב"ד',3),('city','פתח תקווה',4),
  ('city','כפר סבא',5),('city','חולון',6),('city','רחובות',7),('city','לוד',8),
  ('city','ראשון לציון',9),('city','בת ים',10),
  ('grade','גן',1),('grade','א''',2),('grade','ב''',3),('grade','ג''',4),('grade','ד''',5),
  ('grade','ה''',6),('grade','ו''',7),('grade','ז''',8),('grade','ח''',9),
  ('branch','סניף בני ברק',1),('branch','סניף ירושלים',2),('branch','סניף כפר חב"ד',3),
  ('branch','סניף פתח תקווה',4),('branch','סניף כפר סבא',5),('branch','סניף חולון',6),
  ('branch','סניף רחובות',7),('branch','סניף לוד',8),('branch','סניף ראשון לציון',9),
  ('branch','סניף בת ים',10);

-- המנהל הראשון (להחליף למייל שלך)
-- insert into admins (email) values ('your-email@example.com');

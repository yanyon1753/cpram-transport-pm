-- ============================================================
-- CPRAM Fleet - Supabase schema
-- วิธีใช้: เปิดโปรเจกต์ใน supabase.com > SQL Editor > New query
-- แล้ววางไฟล์นี้ทั้งหมด กด Run ครั้งเดียวจบ
-- ============================================================

create extension if not exists "pgcrypto";

create table if not exists vehicles (
  id text primary key,
  brand text not null,
  model text not null,
  year int,
  temp text not null default 'chiller',
  status text not null default 'ready',
  reason text default '',
  last_pm date not null,
  pm_interval int not null default 90,
  driver text default '-',
  created_at timestamptz default now()
);

create table if not exists repairs (
  id uuid primary key default gen_random_uuid(),
  plate text references vehicles(id) on update cascade on delete cascade,
  date date not null,
  type text not null,
  description text not null,
  cost numeric default 0,
  garage text default '-',
  status text default 'เสร็จสิ้น',
  created_at timestamptz default now()
);

create table if not exists drivers (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  license text not null,
  license_type text default 'ท.2',
  expiry date,
  phone text,
  vehicle text default 'สำรอง',
  years int default 0,
  created_at timestamptz default now()
);

alter table vehicles enable row level security;
alter table repairs  enable row level security;
alter table drivers  enable row level security;

drop policy if exists "authenticated full access" on vehicles;
create policy "authenticated full access" on vehicles
  for all to authenticated using (true) with check (true);

drop policy if exists "authenticated full access" on repairs;
create policy "authenticated full access" on repairs
  for all to authenticated using (true) with check (true);

drop policy if exists "authenticated full access" on drivers;
create policy "authenticated full access" on drivers
  for all to authenticated using (true) with check (true);

insert into vehicles (id, brand, model, year, temp, status, reason, last_pm, pm_interval, driver) values
('70-4521','Isuzu','FRR90',2019,'freezer','ready','','2026-04-05',90,'สมชาย ใจดี'),
('70-1188','Isuzu','NPR75',2020,'chiller','ready','','2026-05-01',90,'วิชัย พรหมมา'),
('71-2290','Hino','300',2018,'freezer','not_ready','ซ่อมเครื่องยนต์ - รออะไหล่','2026-03-10',90,'-'),
('70-3345','Isuzu','FRR',2021,'chiller','ready','','2026-06-25',90,'ประยุทธ แสงทอง'),
('71-0087','Hino','500',2017,'freezer','ready','','2026-04-01',90,'อนุชา ศรีสุข'),
('70-5678','Isuzu','ELF',2022,'ambient','ready','','2026-06-10',60,'ธนากร บุญมี'),
('71-1122','Hino','300',2019,'chiller','not_ready','รอตรวจสภาพระบบทำความเย็น','2026-05-20',90,'-'),
('70-9081','Isuzu','FRR',2020,'freezer','ready','','2026-06-29',90,'กิตติ วงศ์ษา'),
('71-3344','Hino','500',2021,'freezer','ready','','2026-04-08',90,'มานพ ทองดี'),
('70-2233','Isuzu','NPR',2018,'chiller','not_ready','รออะไหล่ระบบเบรก','2026-03-01',90,'-'),
('70-6655','Isuzu','ELF',2023,'ambient','ready','','2026-06-15',60,'สุรชัย เพ็ชรดี'),
('71-4499','Hino','300',2020,'freezer','not_ready','อุบัติเหตุ - รอประเมินราคาซ่อม','2026-02-20',90,'-'),
('70-7766','Isuzu','FRR',2022,'chiller','ready','','2026-06-01',90,'เอกชัย รุ่งเรือง')
on conflict (id) do nothing;

insert into repairs (plate, date, type, description, cost, garage, status) values
('71-2290','2026-07-01','เครื่องยนต์','เปลี่ยนปั๊มน้ำมันเชื้อเพลิง',18500,'อู่กลาง CPRAM','กำลังดำเนินการ'),
('71-1122','2026-06-28','ระบบทำความเย็น','ตรวจเช็คคอมเพรสเซอร์ตู้เย็น',9200,'อู่กลาง CPRAM','กำลังดำเนินการ'),
('70-2233','2026-06-30','เบรก','เปลี่ยนผ้าเบรกหน้า-หลัง',6400,'ศูนย์บริการอีซูซุ','กำลังดำเนินการ'),
('71-4499','2026-06-15','ตัวถัง','ซ่อมตัวถังหลังอุบัติเหตุ',42000,'อู่นอก - ช่างเจริญ','กำลังดำเนินการ'),
('70-4521','2026-06-02','ยาง','เปลี่ยนยางคู่หน้า 2 เส้น',7800,'ศูนย์บริการอีซูซุ','เสร็จสิ้น'),
('70-1188','2026-05-18','ระบบไฟฟ้า','ซ่อมไฟท้ายและสายไฟชำรุด',1500,'อู่กลาง CPRAM','เสร็จสิ้น'),
('71-0087','2026-05-10','ระบบทำความเย็น','เติมน้ำยาแอร์ตู้เย็น',3200,'อู่กลาง CPRAM','เสร็จสิ้น'),
('70-3345','2026-05-02','เครื่องยนต์','เปลี่ยนถ่ายน้ำมันเครื่อง + กรอง',2100,'ศูนย์บริการอีซูซุ','เสร็จสิ้น'),
('71-3344','2026-04-22','ตัวถัง','ซ่อมประตูตู้บรรทุกฝืด',2800,'อู่กลาง CPRAM','เสร็จสิ้น'),
('70-9081','2026-04-15','ยาง','สลับยาง ตั้งศูนย์ถ่วงล้อ',1200,'ศูนย์บริการอีซูซุ','เสร็จสิ้น'),
('70-5678','2026-04-06','เครื่องยนต์','ตรวจเช็คระยะ 60,000 กม.',5600,'ศูนย์บริการอีซูซุ','เสร็จสิ้น'),
('70-6655','2026-03-28','ระบบไฟฟ้า','เปลี่ยนแบตเตอรี่',4300,'อู่กลาง CPRAM','เสร็จสิ้น'),
('70-7766','2026-03-20','ระบบทำความเย็น','ล้างคอยล์เย็น ทำความสะอาดระบบ',2600,'อู่กลาง CPRAM','เสร็จสิ้น'),
('71-2290','2026-03-12','เบรก','เปลี่ยนน้ำมันเบรก',900,'อู่กลาง CPRAM','เสร็จสิ้น'),
('70-2233','2026-02-25','เครื่องยนต์','ซ่อมระบบหล่อเย็นเครื่องยนต์',8700,'ศูนย์บริการอีซูซุ','เสร็จสิ้น'),
('71-4499','2026-02-14','ยาง','เปลี่ยนยางทั้งคัน 6 เส้น',21000,'ศูนย์บริการฮีโน่','เสร็จสิ้น'),
('71-0087','2026-01-30','ตัวถัง','ซ่อมสีบริเวณกันชนหน้า',3400,'อู่นอก - ช่างเจริญ','เสร็จสิ้น'),
('70-4521','2026-01-18','ระบบทำความเย็น','เปลี่ยนพัดลมระบายความร้อนตู้เย็น',5200,'อู่กลาง CPRAM','เสร็จสิ้น'),
('70-1188','2026-06-20','เครื่องยนต์','ตรวจเช็คระยะ 80,000 กม.',4900,'ศูนย์บริการอีซูซุ','เสร็จสิ้น'),
('70-9081','2026-06-05','ระบบไฟฟ้า','ซ่อมสวิตช์กระจกไฟฟ้า',800,'อู่กลาง CPRAM','เสร็จสิ้น');

insert into drivers (name, license, license_type, expiry, phone, vehicle, years) values
('สมชาย ใจดี','31-045678-9','ท.2','2027-03-14','081-234-5671','70-4521',8),
('วิชัย พรหมมา','31-118820-2','ท.2','2026-08-02','081-234-5672','70-1188',5),
('ประยุทธ แสงทอง','31-093344-5','ท.3','2027-11-20','081-234-5673','70-3345',11),
('อนุชา ศรีสุข','31-070087-1','ท.2','2026-07-25','081-234-5674','71-0087',3),
('ธนากร บุญมี','31-056781-8','ท.2','2027-01-09','081-234-5675','70-5678',6),
('กิตติ วงศ์ษา','31-090812-6','ท.3','2028-02-17','081-234-5676','70-9081',9),
('มานพ ทองดี','31-133441-3','ท.2','2026-09-30','081-234-5677','71-3344',4),
('สุรชัย เพ็ชรดี','31-066551-7','ท.2','2027-05-05','081-234-5678','70-6655',7),
('เอกชัย รุ่งเรือง','31-077661-4','ท.3','2027-12-01','081-234-5679','70-7766',10),
('ปิยะ ชัยมงคล','31-100223-8','ท.2','2026-10-12','081-234-5680','สำรอง',2);

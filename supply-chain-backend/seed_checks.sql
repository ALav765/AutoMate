-- Run once, after schema.sql: psql supply_plan -f seed_checks.sql
-- Labels, thresholds, and fail_type pulled directly from checklist_engine.py logic.

INSERT INTO checks (check_num, label, description, threshold, fail_type) VALUES
(1,  'HIT Aerosol production band',
     'Monthly HIT Aerosol production should stay within the stated forecast band. Hard ceiling enforced; dipping below the average is flagged but not corrected (forecast-driven).',
     'avg 4.3M units/month, hard ceiling 4.3M × 1.2 = 5.16M', 'calc_error'),

(2,  'Matic Refill production band',
     'Monthly Matic Refill production should stay within the stated forecast band.',
     'avg 0.8M units/month, hard ceiling 1.0M × 1.2 = 1.2M', 'calc_error'),

(3,  'Solid production band',
     'Monthly Solid category production should stay within the stated forecast band.',
     'avg 4.0M units/month, hard ceiling 4.0M × 1.2 = 4.8M', 'calc_error'),

(4,  'Proclin Bleach production band',
     'Monthly Proclin Bleach production should stay within the stated forecast band.',
     'avg 14M units/month, hard ceiling 14M × 1.2 = 16.8M', 'calc_error'),

(5,  'Pocket production band',
     'Monthly Pocket category production should stay within the stated forecast band.',
     'avg 1.8M units/month, hard ceiling 1.8M × 1.2 = 2.16M', 'calc_error'),

(6,  'Aerosol production matches Can & Tins requirement',
     'Can & Tins component requirement should track aerosol-using FG production closely.',
     'ratio (can_req / can-using-FG-production) between 0.97 and 1.05', 'calc_error'),

(7,  'Valve consumption vs aerosol production',
     'Valve component requirement should track HIT Aerosol production specifically (3-category denominator is expected to look low — not an error).',
     'HIT-Aerosol-only ratio between 1.00 and 1.03', 'calc_error'),

(8,  'Perfume & Essential Oils consumption',
     'Monthly perfume/essential oil requirement should stay near the historical average tonnage.',
     '~78 tons/month, flag if outside 74.1–81.9t', 'business_alert'),

(9,  'Vaporer Matic component requirement',
     'Monthly Vaporer Matic component requirement should not exceed the stated cap.',
     '≤ 200,000 units/month', 'business_alert'),

(10, 'Dimefluthrin consumption',
     'Combined Dimefluthrin component requirement (codes 10012285 + 10007043) should not exceed the stated cap.',
     '≤ 2,500 KG/month', 'business_alert'),

(11, 'Glue S04-C consumption',
     'Glue S04-C component requirement should not exceed the stated cap.',
     '≤ 60,000 KG/month', 'business_alert'),

(12, 'Gas Elpiji (TT) requirement',
     'Gas Elpiji component requirement should not exceed the stated cap.',
     '≤ 1,000,000 KG/month', 'business_alert'),

(13, 'HIT Non Stop production matches Printed Carton requirement',
     'Printed Carton / Secondary packaging requirement should track HIT Non Stop FG production.',
     'ratio between 0.97 and 1.05', 'calc_error'),

(14, 'HIT Non Stop production cap',
     'Monthly HIT Non Stop production should not exceed the stated piece cap.',
     '≤ 3,000,000 pieces/month', 'business_alert'),

(15, 'Executive Summary totals match source input',
     'Executive Summary sheet''s monthly forecast totals and opening SOH should exactly match the original input files.',
     'difference < 1 unit for all months and opening SOH', 'calc_error')

ON CONFLICT (check_num) DO NOTHING;

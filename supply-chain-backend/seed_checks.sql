

INSERT INTO checks (check_num, label, description, threshold, fail_type) VALUES
(1, 'Executive Summary totals match source forecast',
    'Executive Summary sheet''s monthly forecast totals should match the original input forecast file.',
    'difference < 1 unit for all months', 'calc_error')
ON CONFLICT (check_num) DO NOTHING;


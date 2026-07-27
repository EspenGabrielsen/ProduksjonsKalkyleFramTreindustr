-- Slett alle rader fra alle tabeller i produksjonskalkyle.db
-- Kjør i SQLite eller via sqlite3 kommandolinje

DELETE FROM products;
DELETE FROM locations;
DELETE FROM work_centers;
DELETE FROM operations;
DELETE FROM item_costs;
DELETE FROM bom_lines;
DELETE FROM routing_lines;
DELETE FROM byproduct_rules;
DELETE FROM capacity_days;
DELETE FROM production_scenarios;
DELETE FROM change_log;
DELETE FROM uploaded_files;
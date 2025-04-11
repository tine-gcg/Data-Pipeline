INSERT INTO "FORMULA" (Unnamed: 0, Unnamed: 1, Unnamed: 2, Unnamed: 3, Unnamed: 4) VALUES ('nan', 'nan', 'nan', 'nan', 'nan');
INSERT INTO "FORMULA" (Unnamed: 0, Unnamed: 1, Unnamed: 2, Unnamed: 3, Unnamed: 4) VALUES ('nan', 'nan', 'nan', 'nan', 'nan');
INSERT INTO "FORMULA" (Unnamed: 0, Unnamed: 1, Unnamed: 2, Unnamed: 3, Unnamed: 4) VALUES ('nan', 'nan', 'nan', 'nan', 'nan');
INSERT INTO "FORMULA" (Unnamed: 0, Unnamed: 1, Unnamed: 2, Unnamed: 3, Unnamed: 4) VALUES ('1.0', '0.0', '0.0', '0.0', '0.0');

SELECT name FROM sqlite_master
WHERE type='table'
AND name NOT LIKE 'sqlite_%';

SELECT * FROM quarterly;
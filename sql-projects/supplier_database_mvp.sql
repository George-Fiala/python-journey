CREATE TABLE suppliers (
    supplier_id INT GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    supplier_name VARCHAR(255) NOT NULL,
    account_number VARCHAR(100),
    supplier_category VARCHAR(100) NOT NULL,
    main_email VARCHAR(255) NOT NULL,
    main_phone VARCHAR(50) NOT NULL,
    is_active BOOLEAN NOT NULL DEFAULT TRUE,
    is_approved BOOLEAN NOT NULL DEFAULT TRUE,
    is_24_7 BOOLEAN NOT NULL DEFAULT FALSE,
    standard_lead_time_days INT,
    default_carriage_cost DECIMAL(10,2),
    notes VARCHAR(1000)
);


INSERT INTO suppliers(
    supplier_name,
    supplier_category,
    main_email,
    main_phone
)
VALUES (
    'Pack Support',
    'Mechanical',
    'contact@packsupport-test.co.uk',
    '+44 0000 000000'
);

SELECT * 
FROM suppliers;
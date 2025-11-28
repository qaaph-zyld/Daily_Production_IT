SELECT 
    t.name AS TableName,
    c.name AS ColumnName,
    ty.name AS DataType,
    c.max_length AS MaxLength,
    c.is_nullable AS IsNullable,
    c.is_identity AS IsIdentity,
    ISNULL(i.is_primary_key, 0) AS IsPrimaryKey
FROM 
    sys.tables t
INNER JOIN 
    sys.columns c ON t.object_id = c.object_id
INNER JOIN 
    sys.types ty ON c.user_type_id = ty.user_type_id
LEFT JOIN 
    (SELECT 
         ic.object_id,
         ic.column_id,
         1 AS is_primary_key
     FROM 
         sys.indexes i
     JOIN 
         sys.index_columns ic ON i.object_id = ic.object_id AND i.index_id = ic.index_id
     WHERE 
         i.is_primary_key = 1) i ON c.object_id = i.object_id AND c.column_id = i.column_id
ORDER BY 
    t.name, c.column_id;
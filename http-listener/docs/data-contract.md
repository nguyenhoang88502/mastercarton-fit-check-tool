# Data Contract - schemas, validation rules and test data

Produced by Worker B (DeepSeek) during iteration 1 and kept as the reference spec for the
listener and the browser UI. `listen.ps1` implements these rules and `tests/test.ps1`
verifies them.

> **One deliberate divergence.** Worker B's INTEGRITY RULES table shortens the orphan-reference
> errors to "FG_ID does not exist." The master prompt specifies the longer
> `FG_ID 'FG-001234' not found in fg_dimensions.csv`, which names the offending value and the
> file to fix. The master prompt wins; the server emits the long form and T12 asserts it.

---
## VALIDATION RULES
```javascript
const VALIDATION = {
  fg_dimensions: {
    fg_id: {
      required: true,
      type: "string",
      pattern: "^[A-Za-z0-9_-]+$",
      maxLength: 50,
      messages: {
        required: "Required",
        pattern: "Must contain only letters, numbers, hyphens, and underscores",
        maxLength: "Must not exceed 50 characters"
      }
    },
    length_mm: {
      required: true,
      type: "number",
      min: 10,
      max: 1000,
      messages: {
        required: "Required",
        type: "Must be a number",
        min: "Must be 10-1000 mm",
        max: "Must be 10-1000 mm"
      }
    },
    width_mm: {
      required: true,
      type: "number",
      min: 10,
      max: 1000,
      messages: {
        required: "Required",
        type: "Must be a number",
        min: "Must be 10-1000 mm",
        max: "Must be 10-1000 mm"
      }
    },
    height_mm: {
      required: true,
      type: "number",
      min: 10,
      max: 1000,
      messages: {
        required: "Required",
        type: "Must be a number",
        min: "Must be 10-1000 mm",
        max: "Must be 10-1000 mm"
      }
    },
    target_qty: {
      required: false,
      type: "integer",
      min: 1,
      max: 10000,
      messages: {
        type: "Must be an integer",
        min: "Must be 1-10000",
        max: "Must be 1-10000"
      }
    },
    source: {
      required: false,
      type: "enum",
      enum: ["MANUAL_ENTRY", "D365_IMPORT", "MOBILE_APP", "BATCH_IMPORT"],
      default: "MANUAL_ENTRY",
      messages: {
        enum: "Must be one of: MANUAL_ENTRY, D365_IMPORT, MOBILE_APP, BATCH_IMPORT"
      }
    }
  },
  carton_dimensions: {
    carton_id: {
      required: true,
      type: "string",
      pattern: "^[A-Za-z0-9_-]+$",
      maxLength: 50,
      messages: {
        required: "Required",
        pattern: "Must contain only letters, numbers, hyphens, and underscores",
        maxLength: "Must not exceed 50 characters"
      }
    },
    length_mm: {
      required: true,
      type: "number",
      min: 10,
      max: 1000,
      messages: {
        required: "Required",
        type: "Must be a number",
        min: "Must be 10-1000 mm",
        max: "Must be 10-1000 mm"
      }
    },
    width_mm: {
      required: true,
      type: "number",
      min: 10,
      max: 1000,
      messages: {
        required: "Required",
        type: "Must be a number",
        min: "Must be 10-1000 mm",
        max: "Must be 10-1000 mm"
      }
    },
    height_mm: {
      required: true,
      type: "number",
      min: 10,
      max: 1000,
      messages: {
        required: "Required",
        type: "Must be a number",
        min: "Must be 10-1000 mm",
        max: "Must be 10-1000 mm"
      }
    },
    dimension_source: {
      required: true,
      type: "enum",
      enum: ["BOM_LINKED_CA_AVERAGE", "MANUAL_MEASUREMENT", "SUPPLIER_SPEC"],
      messages: {
        required: "Required",
        enum: "Must be one of: BOM_LINKED_CA_AVERAGE, MANUAL_MEASUREMENT, SUPPLIER_SPEC"
      }
    }
  },
  bom_carton_assignments: {
    fg_id: {
      required: true,
      type: "string",
      pattern: "^[A-Za-z0-9_-]+$",
      maxLength: 50,
      messages: {
        required: "Required",
        pattern: "Must contain only letters, numbers, hyphens, and underscores",
        maxLength: "Must not exceed 50 characters"
      }
    },
    carton_id: {
      required: true,
      type: "string",
      pattern: "^[A-Za-z0-9_-]+$",
      maxLength: 50,
      messages: {
        required: "Required",
        pattern: "Must contain only letters, numbers, hyphens, and underscores",
        maxLength: "Must not exceed 50 characters"
      }
    },
    source_bom_date: {
      required: true,
      type: "date",
      pattern: "^\\d{4}-\\d{2}-\\d{2}$",
      messages: {
        required: "Required",
        pattern: "Must be a valid date format YYYY-MM-DD",
        date: "Must be a real calendar date"
      }
    },
    confidence: {
      required: false,
      type: "enum",
      enum: ["HIGH", "MEDIUM", "LOW"],
      default: "MEDIUM",
      messages: {
        enum: "Must be one of: HIGH, MEDIUM, LOW"
      }
    }
  }
};
```

## INTEGRITY RULES
| Rule Name                     | Endpoint                   | Check                                                           | HTTP Status | Error String                                               |
|-------------------------------|----------------------------|-----------------------------------------------------------------|-------------|------------------------------------------------------------|
| Duplicate FG ID               | fg_dimensions              | FG_ID already exists in CSV                                     | 409         | "FG_ID already exists."                                    |
| Duplicate Carton ID           | carton_dimensions          | Carton_ID already exists in CSV                                 | 409         | "Carton_ID already exists."                                |
| Duplicate Assignment          | bom_carton_assignments     | Composite (FG_ID, Carton_ID) already exists in CSV              | 409         | "Assignment for this FG_ID and Carton_ID already exists."  |
| Orphan FG ID                  | bom_carton_assignments     | FG_ID does not exist in fg_dimensions.csv                       | 400         | "FG_ID does not exist."                                    |
| Orphan Carton ID              | bom_carton_assignments     | Carton_ID does not exist in carton_dimensions.csv               | 400         | "Carton_ID does not exist."                                |

## TEST DATA
```json
[
  {
    "fg_id": "FG-1000001",
    "length_mm": 180,
    "width_mm": 60,
    "height_mm": 45,
    "target_qty": 500,
    "source": "D365_IMPORT"
  },
  {
    "fg_id": "FG-1000002",
    "length_mm": 150,
    "width_mm": 50,
    "height_mm": 40,
    "target_qty": 300,
    "source": "MOBILE_APP"
  },
  {
    "fg_id": "FG-1000003",
    "length_mm": 160,
    "width_mm": 55,
    "height_mm": 50,
    "source": "MANUAL_ENTRY"
  },
  {
    "fg_id": "FG-1000004",
    "length_mm": 200,
    "width_mm": 100,
    "height_mm": 70,
    "target_qty": 200,
    "source": "BATCH_IMPORT"
  },
  {
    "fg_id": "FG-1000005",
    "length_mm": 190,
    "width_mm": 80,
    "height_mm": 55,
    "target_qty": 600,
    "source": "MANUAL_ENTRY"
  },
  {
    "fg_id": "FG-1000006",
    "length_mm": 140,
    "width_mm": 45,
    "height_mm": 35,
    "target_qty": 1000,
    "source": "D365_IMPORT"
  },
  {
    "fg_id": "FG-1000007",
    "length_mm": 175,
    "width_mm": 70,
    "height_mm": 50,
    "target_qty": 150,
    "source": "MOBILE_APP"
  },
  {
    "fg_id": "FG-1000008",
    "length_mm": 210,
    "width_mm": 110,
    "height_mm": 75,
    "target_qty": 80,
    "source": "BATCH_IMPORT"
  },
  {
    "fg_id": "FG-1000009",
    "length_mm": 155,
    "width_mm": 55,
    "height_mm": 42,
    "target_qty": 400,
    "source": "D365_IMPORT"
  },
  {
    "fg_id": "FG-1000010",
    "length_mm": 165,
    "width_mm": 60,
    "height_mm": 45,
    "target_qty": 250,
    "source": "MANUAL_ENTRY"
  }
]
```
```json
[
  {
    "carton_id": "4000001",
    "length_mm": 300,
    "width_mm": 200,
    "height_mm": 150,
    "dimension_source": "SUPPLIER_SPEC"
  },
  {
    "carton_id": "4000002",
    "length_mm": 250,
    "width_mm": 180,
    "height_mm": 120,
    "dimension_source": "MANUAL_MEASUREMENT"
  },
  {
    "carton_id": "4000003",
    "length_mm": 350,
    "width_mm": 220,
    "height_mm": 160,
    "dimension_source": "BOM_LINKED_CA_AVERAGE"
  },
  {
    "carton_id": "4000004",
    "length_mm": 280,
    "width_mm": 190,
    "height_mm": 130,
    "dimension_source": "SUPPLIER_SPEC"
  },
  {
    "carton_id": "4000005",
    "length_mm": 320,
    "width_mm": 210,
    "height_mm": 140,
    "dimension_source": "MANUAL_MEASUREMENT"
  },
  {
    "carton_id": "4000006",
    "length_mm": 260,
    "width_mm": 170,
    "height_mm": 110,
    "dimension_source": "BOM_LINKED_CA_AVERAGE"
  },
  {
    "carton_id": "4000007",
    "length_mm": 290,
    "width_mm": 200,
    "height_mm": 150,
    "dimension_source": "SUPPLIER_SPEC"
  },
  {
    "carton_id": "4000008",
    "length_mm": 310,
    "width_mm": 220,
    "height_mm": 160,
    "dimension_source": "MANUAL_MEASUREMENT"
  },
  {
    "carton_id": "4000009",
    "length_mm": 270,
    "width_mm": 180,
    "height_mm": 120,
    "dimension_source": "BOM_LINKED_CA_AVERAGE"
  },
  {
    "carton_id": "4000010",
    "length_mm": 340,
    "width_mm": 230,
    "height_mm": 170,
    "dimension_source": "SUPPLIER_SPEC"
  }
]
```
```json
[
  {
    "fg_id": "FG-1000001",
    "carton_id": "4000001",
    "source_bom_date": "2025-02-15",
    "confidence": "HIGH"
  },
  {
    "fg_id": "FG-1000002",
    "carton_id": "4000002",
    "source_bom_date": "2025-03-01"
  },
  {
    "fg_id": "FG-1000003",
    "carton_id": "4000003",
    "source_bom_date": "2025-03-10",
    "confidence": "LOW"
  },
  {
    "fg_id": "FG-1000004",
    "carton_id": "4000004",
    "source_bom_date": "2025-01-20",
    "confidence": "MEDIUM"
  },
  {
    "fg_id": "FG-1000005",
    "carton_id": "4000005",
    "source_bom_date": "2025-04-05"
  },
  {
    "fg_id": "FG-1000006",
    "carton_id": "4000006",
    "source_bom_date": "2025-02-28",
    "confidence": "HIGH"
  },
  {
    "fg_id": "FG-1000007",
    "carton_id": "4000007",
    "source_bom_date": "2025-03-18"
  },
  {
    "fg_id": "FG-1000008",
    "carton_id": "4000008",
    "source_bom_date": "2025-02-01",
    "confidence": "LOW"
  },
  {
    "fg_id": "FG-1000009",
    "carton_id": "4000009",
    "source_bom_date": "2025-05-12",
    "confidence": "MEDIUM"
  },
  {
    "fg_id": "FG-1000010",
    "carton_id": "4000010",
    "source_bom_date": "2025-01-05"
  }
]
```

## EDGE CASES
```json
[
  {
    "_expect": {
      "endpoint": "fg_dimensions",
      "status": 400,
      "field": "length_mm",
      "message": "Required"
    },
    "fg_id": "FG-2000001",
    "width_mm": 50,
    "height_mm": 50
  },
  {
    "_expect": {
      "endpoint": "fg_dimensions",
      "status": 400,
      "field": "length_mm",
      "message": "Must be 10-1000 mm"
    },
    "fg_id": "FG-2000002",
    "length_mm": 5,
    "width_mm": 50,
    "height_mm": 50,
    "source": "MANUAL_ENTRY"
  },
  {
    "_expect": {
      "endpoint": "fg_dimensions",
      "status": 400,
      "field": "length_mm",
      "message": "Must be 10-1000 mm"
    },
    "fg_id": "FG-2000003",
    "length_mm": 1500,
    "width_mm": 50,
    "height_mm": 50,
    "source": "D365_IMPORT"
  },
  {
    "_expect": {
      "endpoint": "fg_dimensions",
      "status": 400,
      "field": "length_mm",
      "message": "Must be a number"
    },
    "fg_id": "FG-2000004",
    "length_mm": "abc",
    "width_mm": 50,
    "height_mm": 50,
    "source": "MOBILE_APP"
  },
  {
    "_expect": {
      "endpoint": "fg_dimensions",
      "status": 400,
      "field": "fg_id",
      "message": "Must not exceed 50 characters"
    },
    "fg_id": "A123456789012345678901234567890123456789012345678901",
    "length_mm": 100,
    "width_mm": 80,
    "height_mm": 60
  },
  {
    "_expect": {
      "endpoint": "fg_dimensions",
      "status": 400,
      "field": "fg_id",
      "message": "Must contain only letters, numbers, hyphens, and underscores"
    },
    "fg_id": "FG-123$%",
    "length_mm": 100,
    "width_mm": 80,
    "height_mm": 60
  },
  {
    "_expect": {
      "endpoint": "fg_dimensions",
      "status": 400,
      "field": "source",
      "message": "Must be one of: MANUAL_ENTRY, D365_IMPORT, MOBILE_APP, BATCH_IMPORT"
    },
    "fg_id": "FG-2000005",
    "length_mm": 100,
    "width_mm": 80,
    "height_mm": 60,
    "source": "INVALID"
  },
  {
    "_expect": {
      "endpoint": "bom_carton_assignments",
      "status": 400,
      "field": "source_bom_date",
      "message": "Must be a valid date format YYYY-MM-DD"
    },
    "fg_id": "FG-1000001",
    "carton_id": "4000001",
    "source_bom_date": "2025-13-01"
  },
  {
    "_expect": {
      "endpoint": "bom_carton_assignments",
      "status": 400,
      "field": "source_bom_date",
      "message": "Must be a real calendar date"
    },
    "fg_id": "FG-1000002",
    "carton_id": "4000002",
    "source_bom_date": "2026-02-30"
  },
  {
    "_expect": {
      "endpoint": "fg_dimensions",
      "status": 400,
      "field": "target_qty",
      "message": "Must be 1-10000"
    },
    "fg_id": "FG-2000006",
    "length_mm": 120,
    "width_mm": 70,
    "height_mm": 50,
    "target_qty": 0,
    "source": "BATCH_IMPORT"
  },
  {
    "_expect": {
      "endpoint": "fg_dimensions",
      "status": 400,
      "field": "target_qty",
      "message": "Must be 1-10000"
    },
    "fg_id": "FG-2000007",
    "length_mm": 120,
    "width_mm": 70,
    "height_mm": 50,
    "target_qty": 10001,
    "source": "MANUAL_ENTRY"
  },
  {
    "_expect": {
      "endpoint": "bom_carton_assignments",
      "status": 400,
      "field": "fg_id",
      "message": "FG_ID does not exist."
    },
    "fg_id": "FG-9999999",
    "carton_id": "4000003",
    "source_bom_date": "2025-06-01"
  }
]
```

## NOTES
- The `Timestamp` column is automatically generated by the server on append; clients must never send it, and it is excluded from these validation rules.
- Field validations (type, required, pattern, range, enum, date format) run before referentialâintegrity checks. Only fields that pass syntax/semantic validation reach the duplicateâkey and foreignâkey checks.
- For `source_bom_date`, the server must validate the string against `YYYY-MM-DD` format and then confirm it represents a real calendar date (e.g., using `DateTime.TryParseExact` with a real calendar check) to catch dates like 2026â02â30.
- All dimensions are in millimetres. The numeric range `10â1000` is inclusive at both ends.
- When `source` or `confidence` is omitted from a request, the server applies the default value before writing the row.

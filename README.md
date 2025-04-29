import pandas as pd

# Load the Excel file
excel_path = 'your_file.xlsx'

# Load specific sheets by name
new_df = pd.read_excel(excel_path, sheet_name='Plan')      # your new sheet
old_df = pd.read_excel(excel_path, sheet_name='Plan Old')  # your old sheet


import pandas as pd
from datetime import datetime

# Load both sheets
new_df = pd.read_excel("new_sheet.xlsx")
old_df = pd.read_excel("old_sheet.xlsx")

# Ensure 'Recovery Plan' columns are datetime
new_df['Recovery Plan'] = pd.to_datetime(new_df['Recovery Plan'], errors='coerce')
old_df['Recovery Plan'] = pd.to_datetime(old_df['Recovery Plan'], errors='coerce')

# Merge on 'Code' and 'DR Scenerio'
merged = pd.merge(new_df, old_df, on=['Code', 'DR Scenerio'], suffixes=('_new', '_old'), how='left')

# Define a function to apply row-wise logic
def resolve_row(row):
    # Extract relevant fields
    rp_new = row['Recovery Plan_new']
    rp_old = row['Recovery Plan_old']

    # If no match found in old (app doesn't exist)
    if pd.isna(row['RTE_old']) and pd.isna(row['RPE_old']):
        return pd.Series({
            'RTE': row['RTE_new'],
            'RPE': row['RPE_new'],
            'Recovery Plan': rp_new
        })

    # If both are blank
    if pd.isna(rp_new) and pd.isna(rp_old):
        return pd.Series({
            'RTE': row['RTE_new'],
            'RPE': row['RPE_new'],
            'Recovery Plan': None
        })

    # If only new has date
    if not pd.isna(rp_new) and pd.isna(rp_old):
        return pd.Series({
            'RTE': row['RTE_new'],
            'RPE': row['RPE_new'],
            'Recovery Plan': rp_new
        })

    # If only old has date
    if pd.isna(rp_new) and not pd.isna(rp_old):
        return pd.Series({
            'RTE': row['RTE_old'],
            'RPE': row['RPE_old'],
            'Recovery Plan': rp_old
        })

    # If both have date, use latest
    if rp_new >= rp_old:
        return pd.Series({
            'RTE': row['RTE_new'],
            'RPE': row['RPE_new'],
            'Recovery Plan': rp_new
        })
    else:
        return pd.Series({
            'RTE': row['RTE_old'],
            'RPE': row['RPE_old'],
            'Recovery Plan': rp_old
        })

# Apply the logic row by row
resolved = merged.apply(resolve_row, axis=1)

# Append resolved values to the original new_df (keeping its other columns)
final_df = new_df.copy()
final_df['RTE'] = resolved['RTE']
final_df['RPE'] = resolved['RPE']
final_df['Recovery Plan'] = resolved['Recovery Plan']

# Export to Excel
final_df.to_excel("merged_output.xlsx", index=False)






# excel-codes
=IF(
  ISNA(MATCH(1, (PLAN_OLD!B:B=B2)*(PLAN_OLD!E:E=E2), 0)),
  G2,
  IF(
    AND(I2="", INDEX(PLAN_OLD!I:I, MATCH(1, (PLAN_OLD!B:B=B2)*(PLAN_OLD!E:E=E2), 0))=""),
    G2,
    IF(
      I2="",
      INDEX(PLAN_OLD!G:G, MATCH(1, (PLAN_OLD!B:B=B2)*(PLAN_OLD!E:E=E2), 0)),
      IF(
        INDEX(PLAN_OLD!I:I, MATCH(1, (PLAN_OLD!B:B=B2)*(PLAN_OLD!E:E=E2), 0))="",
        G2,
        IF(
          INDEX(PLAN_OLD!I:I, MATCH(1, (PLAN_OLD!B:B=B2)*(PLAN_OLD!E:E=E2), 0)) > I2,
          INDEX(PLAN_OLD!G:G, MATCH(1, (PLAN_OLD!B:B=B2)*(PLAN_OLD!E:E=E2), 0)),
          G2
        )
      )
    )
  )
)


=IF(
  AND(I2="", INDEX(PLAN_OLD!I:I, MATCH(1, (PLAN_OLD!B:B=B2)*(PLAN_OLD!E:E=E2), 0))=""),
  G2,
  IF(
    I2="",
    INDEX(PLAN_OLD!G:G, MATCH(1, (PLAN_OLD!B:B=B2)*(PLAN_OLD!E:E=E2), 0)),
    IF(
      INDEX(PLAN_OLD!I:I, MATCH(1, (PLAN_OLD!B:B=B2)*(PLAN_OLD!E:E=E2), 0))="",
      G2,
      IF(
        INDEX(PLAN_OLD!I:I, MATCH(1, (PLAN_OLD!B:B=B2)*(PLAN_OLD!E:E=E2), 0)) > I2,
        INDEX(PLAN_OLD!G:G, MATCH(1, (PLAN_OLD!B:B=B2)*(PLAN_OLD!E:E=E2), 0)),
        G2
      )
    )
  )
)




=IF(
  AND(I2="", INDEX(PLAN_OLD!I:I, MATCH(1, (PLAN_OLD!B:B=B2)*(PLAN_OLD!E:E=E2), 0))=""),
  J2,
  IF(
    I2="",
    INDEX(PLAN_OLD!J:J, MATCH(1, (PLAN_OLD!B:B=B2)*(PLAN_OLD!E:E=E2), 0)),
    IF(
      INDEX(PLAN_OLD!I:I, MATCH(1, (PLAN_OLD!B:B=B2)*(PLAN_OLD!E:E=E2), 0))="",
      J2,
      IF(
        INDEX(PLAN_OLD!I:I, MATCH(1, (PLAN_OLD!B:B=B2)*(PLAN_OLD!E:E=E2), 0)) > I2,
        INDEX(PLAN_OLD!J:J, MATCH(1, (PLAN_OLD!B:B=B2)*(PLAN_OLD!E:E=E2), 0)),
        J2
      )
    )
  )
)


=IF(
  AND(I2="", INDEX(Sheet2!I:I, MATCH(1, (Sheet2!B:B=B2)*(Sheet2!E:E=E2), 0))=""),
  "",
  IF(
    I2="",
    INDEX(Sheet2!I:I, MATCH(1, (Sheet2!B:B=B2)*(Sheet2!E:E=E2), 0)),
    IF(
      INDEX(Sheet2!I:I, MATCH(1, (Sheet2!B:B=B2)*(Sheet2!E:E=E2), 0))="",
      I2,
      IF(
        INDEX(Sheet2!I:I, MATCH(1, (Sheet2!B:B=B2)*(Sheet2!E:E=E2), 0)) > I2,
        INDEX(Sheet2!I:I, MATCH(1, (Sheet2!B:B=B2)*(Sheet2!E:E=E2), 0)),
        I2
      )
    )
  )
)


=IF(
  OR(J2="", INDEX(Sheet2!J:J, MATCH(1, (Sheet2!B:B=B2)*(Sheet2!C:C=C2), 0))=""),
  IF(J2="", INDEX(Sheet2!U:U, MATCH(1, (Sheet2!B:B=B2)*(Sheet2!C:C=C2), 0)), M2),
  IF(
    INDEX(Sheet2!J:J, MATCH(1, (Sheet2!B:B=B2)*(Sheet2!C:C=C2), 0)) > J2,
    INDEX(Sheet2!U:U, MATCH(1, (Sheet2!B:B=B2)*(Sheet2!C:C=C2), 0)),
    M2
  )
)

=IF(
  OR(I2="", INDEX(Sheet2!I:I, MATCH(1, (Sheet2!B:B=B2)*(Sheet2!E:E=E2), 0))=""),
  IF(I2="", INDEX(Sheet2!I:I, MATCH(1, (Sheet2!B:B=B2)*(Sheet2!E:E=E2), 0)), I2),
  IF(
    INDEX(Sheet2!I:I, MATCH(1, (Sheet2!B:B=B2)*(Sheet2!E:E=E2), 0)) > I2,
    INDEX(Sheet2!I:I, MATCH(1, (Sheet2!B:B=B2)*(Sheet2!E:E=E2), 0)),
    I2
  )
)



=IFERROR(
  IF(
    INDEX(Sheet2!J:J, MATCH(1, (Sheet2!B:B=B2)*(Sheet2!C:C=C2), 0)) = "",
    IF(M2="", "", M2),
    IF(
      INDEX(Sheet2!J:J, MATCH(1, (Sheet2!B:B=B2)*(Sheet2!C:C=C2), 0)) > J2,
      INDEX(Sheet2!U:U, MATCH(1, (Sheet2!B:B=B2)*(Sheet2!C:C=C2), 0)),
      IF(M2="", "", M2)
    )
  ),
  ""
)


=IFERROR(
  IF(
    INDEX(Sheet2!J:J, MATCH(1, (Sheet2!B:B=B2)*(Sheet2!C:C=C2), 0)) > J2,
    INDEX(Sheet2!J:J, MATCH(1, (Sheet2!B:B=B2)*(Sheet2!C:C=C2), 0)),
    J2
  ),
  J2
)


=IF(
  OR(J2="", XLOOKUP(B2, Sheet2!B:B, Sheet2!J:J, "") = ""),
  IF(J2="", XLOOKUP(B2, Sheet2!B:B, Sheet2!U:U, ""), M2),
  IF(
    XLOOKUP(B2, Sheet2!B:B, Sheet2!J:J, "") > J2,
    XLOOKUP(B2, Sheet2!B:B, Sheet2!U:U, ""),
    M2
  )
)

=LET(
  appID, B2,
  drScenion, C2,
  rev1, J2,
  rev2, XLOOKUP(1, (Sheet2!B:B=appID)*(Sheet2!C:C=drScenion), Sheet2!J:J, ""),
  rto2, XLOOKUP(1, (Sheet2!B:B=appID)*(Sheet2!C:C=drScenion), Sheet2!U:U, ""),
  IF(
    OR(rev1="", rev2=""),
    IF(rev1="", rto2, M2),
    IF(rev2 > rev1, rto2, M2)
  )
)


# excel-codes
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


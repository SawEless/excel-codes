# excel-codes
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


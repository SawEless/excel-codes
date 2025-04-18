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

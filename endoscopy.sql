SELECT
	/* TCI pulling list for Endo patients */
	pt.FORENAME + ' ' + pt.SURNAME 'Patient Name'
	, item.OFFERDTTM
	, ser.NAME
	, addr.LINE1 'mobile'
FROM APT.APENTRY apentry
JOIN PIM.PATIENT pt
	ON apentry.PATIENTOID = pt.OID
	AND pt.STATUS = 'A'
JOIN PIM.PATIENTADDRESSROLE ptaddrole_msgmob
	ON ptaddrole_msgmob.IDENTIFYINGTYPE = 'Patient'
	AND pt.OID = ptaddrole_msgmob.IDENTIFYINGOID
	AND ptaddrole_msgmob.ROTYPCODE = 'CC_MSGMOB'
	AND ptaddrole_msgmob.STATUS = 'A'
	AND (ptaddrole_msgmob.ENDDTTM = convert(datetime, 0)
		OR (
			ptaddrole_msgmob.STARTDTTM < convert(date, GETDATE())
			AND ptaddrole_msgmob.ENDDTTM > convert(date, GETDATE())
			)
		)
JOIN PIM.PATIENTADDRESS addr
	ON ptaddrole_msgmob.ADDRESSOID = addr.OID
	AND addr.STATUS = 'A'
JOIN APT.APEOFFER offer
	ON apentry.OID = offer.APENTRYOID
	AND offer.STATUS = 'A'
JOIN APT.APEOFFERITEM item
	ON offer.OID = item.APEOFFEROID
	AND item.IDENTIFYINGTYPE = 'CPMBOOKING' /* inpatient booking */
	AND item.STATUS = 'A'
JOIN CAP.CPMBOOKING book
	ON item.IDENTIFYINGOID = book.OID
	AND book.BKSTACODE != 'C' /* Not a cancelled offer */
	AND book.STATUS = 'A'
JOIN ENC.PATIENTAPPOINTMENTIP apptip
	on book.OID = apptip.CPMBOOKINGOID
	AND apptip.STATUS = 'A'
JOIN EMM.SERVICE ser
	ON apptip.SERVICEOID = ser.OID	
	AND ser.STATUS = 'A'
WHERE 
	CONVERT(date, item.OFFERDTTM) = DATEADD(day, 0, CONVERT(date, getdate()))
	AND ser.OID = '3900000820' /* L3 Gate 13 Endo */
ORDER BY item.OFFERDTTM
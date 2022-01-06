::[Bat To Exe Converter]
::
::YAwzoRdxOk+EWAnk
::fBw5plQjdG8=
::YAwzuBVtJxjWCl3EqQJgSA==
::ZR4luwNxJguZRRnk
::Yhs/ulQjdF+5
::cxAkpRVqdFKZSDk=
::cBs/ulQjdF+5
::ZR41oxFsdFKZSTk=
::eBoioBt6dFKZSDk=
::cRo6pxp7LAbNWATEpCI=
::egkzugNsPRvcWATEpCI=
::dAsiuh18IRvcCxnZtBJQ
::cRYluBh/LU+EWAnk
::YxY4rhs+aU+IeA==
::cxY6rQJ7JhzQF1fEqQJlZksaHErSXA==
::ZQ05rAF9IBncCkqN+0xwdVsBAlTMbCXqZg==
::ZQ05rAF9IAHYFVzEqQIUMT5aT1G9Hn6zCrE50M3EzOWVpy0=
::eg0/rx1wNQPfEVWB+kM9LVsJDCCbGWW5U4o+/eH368+/h3I+W/A6GA==
::fBEirQZwNQPfEVWB+kM9LVsJDCCbGWW5U4o+/eH368+/h3I+W/A6GA==
::cRolqwZ3JBvQF1fEqQIUMT5aT1EK26EywXaZBgN9GQ0DeZuaoB25kH9eI3zu
::dhA7uBVwLU+EWNBZBMviyZ6DiYBaxu5+wTDRLqhzMiUDeZuYgye4sGReBlQLzxFCbaopNRDQKv+Uf6s=
::YQ03rBFzNR3SWATElA==
::dhAmsQZ3MwfNWATEfZ/Aocm1ydNHuLNAg3HbbSU9r2JabcnGgmfovQgTMTFSwGX8GxRjmlttUIemHQXrbA==
::ZQ0/vhVqMQ3MEVWAtB9wSA==
::Zg8zqx1/OA3MEVWAtB9wSA==
::dhA7pRFwIByZRRnk
::Zh4grVQjdCyDJGyX8VAjFDpQQQ2MNXiuFLQI5/rHy++UqVkSRN65jl9eAW4I1geXMZNhsU9tni8TpO8VKRVbKy2JewY4rUt6k1umONWZ/Qr5Tyg=
::YB416Ek+ZG8=
::
::
::978f952a14a936cc963da21a135fa983
set OUT=%1
set IN=%2
set SRC=%3
set ICN=%4
rar a -ep %OUT% -sfx DyEncGUI5.0.OtherSettings.config DyEncryptor5.0_Plus.exe Dy_EncCore.exe DyEncGUI5.0.config DyEnc_FileDestroyModule.exe DyEncIcon.ico GUI_Color.config DyEnc5.0.HISTORY %IN% 
rar c -z%SRC% %OUT%
WinRAR s -iicon%ICN% %OUT%
del /F /Q %SRC%
-- update result preparestatment
-- 1 filename 2 filedata 3 caseNO 4 revision
update CaseRevision set ReortFileName = ?,ReortFileData = ? where CASENO = ? and REVISION = ?
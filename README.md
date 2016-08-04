# TrustMe
A quick and easy VBA exploit to change Office Trust Centre settings in the registry without user knowledge

To use, import into a microsoft office document. Then, simply modify the virusCode string in the bas file to whatever you please. Provide a trigger for the TrustMe subroutine (I prefer an On Workbook Open event). If the user allows the Macro to run, the VBProject model will be yours to use and manipulate.

Note: This should permanently change trust centre settings until the user realizes what's going on. You might consider changing them back after your code has executed if you intend to be sneaky.

ANOTHER NOTE: This is not intended to be used for illegal purposes. It's just good fun.

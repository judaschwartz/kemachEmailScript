# kemachEmailScript

Use this script to genrate emails for orders from a google form like [this](https://docs.google.com/forms/d/1R0lu2e-WaseXk69aCQEdmrqJUQfQm3A9ksUrhcQp3zU/edit) (please do not edit the sample form)
the second row of the results sheet should conatian the names of the fields and the third row the prices
the last columns of the sheet should be named in this order `Number of Items`	`Order ID`	`total` `edit link`
on pesach there is a coupons column and two more columns before the `total` column named `subtotal` and `Coupon discount`
This script should be pasted to the "apps script" which can be accessed from the extensions menu of the forms results spreadsheet.
Then from the triggers menu create a new trigger to run `triggerOnSubmi`t and Select event type "On form submit"

The `processOrderData` and `processPaymentData` functions create fixed length result files to be fed to the cobalt system

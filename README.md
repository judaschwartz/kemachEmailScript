# kemachEmailScript

Use this script to genrate emails for orders from a google form like [this](https://forms.gle/5L9mvcz8vApmgada6)
the second row of the results sheet should conatian the names of the fields and the third row the prices
the last columns of the sheet should be named in this order `Number of Items`	`Order ID`	`total` `edit link`
on pesach there is a coupons column and two more columns before the `total` column named `subtotal` `Coupon discount`

The `processOrderData` and `processPaymentData` functions create fixed length result files to be fed to the cobalt system

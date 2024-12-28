# Copy email addresses from brightspace

There is no easy way (that I know) to extract the email addresses of students in brightspace.

So, here is a workaround. Open the mail classlist dialog, and paste this in the devtools
It filters the students (annoyingly BS always adds the teachers) and provides a neat list that can be copy pasted into outlook.

```
[...document.querySelectorAll(".d2l-multiselect-choice")].filter(n => n.innerHTML.includes('student')).map(n => n.querySelector('span').textContent).join(';') 
```
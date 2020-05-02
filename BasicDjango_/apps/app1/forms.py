from django import forms

class SenderForm(forms.Form):
    LOGIN = forms.EmailField(max_length=30)
    PWD = forms.CharField(widget = forms.PasswordInput())
    WHOM = forms.FileField()
    ATTACH_TPL = forms.FileField(required = False)
    THEME = forms.CharField( max_length=50)
    LETTER = forms.CharField(widget = forms.Textarea(attrs={'cols':30,'rows':10}))
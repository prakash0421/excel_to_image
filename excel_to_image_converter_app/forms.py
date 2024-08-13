from django import forms

class UploadFileForm(forms.Form):
    file = forms.FileField(
        label='Select Excel File',
        widget=forms.ClearableFileInput(attrs={'class': 'form-control'})
    )
    email_to = forms.EmailField(
        label='Recipient Email',
        widget=forms.EmailInput(attrs={'class': 'form-control'})
    )
    user_name = forms.CharField(
        label='Your Name',
        max_length=100,
        widget=forms.TextInput(attrs={'class': 'form-control'})
    )
    email_body = forms.CharField(
        label='Email Body',
        widget=forms.Textarea(attrs={'class': 'form-control custom-textarea'}),
        required=False  # Body is optional
    )
    image_type = forms.ChoiceField(
        choices=[('full', 'Full Image'), ('specific', 'Filtered Image')],
        label='Image Type',
        widget=forms.Select(attrs={'class': 'form-control custom-select'})
    )

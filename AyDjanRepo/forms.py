
from django.contrib.auth.models import User
from django import forms
from django.contrib.auth.forms import UserCreationForm


class CustomUserCreationForm(UserCreationForm):
    email = forms.EmailField(
        max_length=254, help_text='Required. Enter a valid email address.')

    class Meta:
        model = User
        fields = ('email', 'password1', 'password2')

    def clean_email(self):
        email = self.cleaned_data.get('email')
        allowed_domains = ['lbaik.com'',2coom.com',
                           'devo-p.com', 'aljouai.com']
        if not any(email.endswith(domain) for domain in allowed_domains):
            raise forms.ValidationError(
                "Email must have an allowed domain (e.g., gmail.com, yahoo.com)")

        return email

from django.test import TestCase, Client
from django.urls import reverse
from users.models import CustomUser
import datetime

class NotasListViewRedirectTest(TestCase):
    def setUp(self):
        # Create a user to authenticate
        self.user = CustomUser.objects.create_user(
            username='testuser',
            password='testpassword'
        )
        self.client = Client()

    def test_redirects_without_month(self):
        # Force login
        self.client.login(username='testuser', password='testpassword')
        
        # Access the list view without data_emissao_month
        url = reverse('lista_notas')
        response = self.client.get(url)
        
        # Verify redirect
        hoje = datetime.date.today().strftime("%Y-%m")
        expected_url = f"{url}?data_emissao_month={hoje}"
        self.assertRedirects(response, expected_url)

    def test_no_redirect_with_month(self):
        # Force login
        self.client.login(username='testuser', password='testpassword')
        
        # Access the list view with data_emissao_month
        url = reverse('lista_notas')
        response = self.client.get(url, {'data_emissao_month': '2026-05'})
        
        # Verify no redirect (status 200)
        self.assertEqual(response.status_code, 200)

    def test_preserves_other_params(self):
        # Force login
        self.client.login(username='testuser', password='testpassword')
        
        # Access the list view with a filial parameter but without data_emissao_month
        url = reverse('lista_notas')
        response = self.client.get(url, {'filial': '123'})
        
        # Verify redirect preserves the filial parameter
        hoje = datetime.date.today().strftime("%Y-%m")
        self.assertEqual(response.status_code, 302)
        redirect_target = response.url
        self.assertIn('filial=123', redirect_target)
        self.assertIn(f'data_emissao_month={hoje}', redirect_target)

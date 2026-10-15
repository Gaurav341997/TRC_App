from django.contrib import admin
from django.urls import include, path


urlpatterns = [
    # Django Admin
    path("admin/", admin.site.urls),

    # Standard Django authentication
    path("accounts/", include("django.contrib.auth.urls")),

    # SAML SSO endpoints
    path("saml2/", include("djangosaml2.urls")),

    # TRC application
    path("", include("rates.urls")),
]
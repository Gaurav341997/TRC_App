from django.urls import path

from . import views


urlpatterns = [
    # Application Home Page
    path("", views.home, name="home"),

    # Rate Console Dashboard
    path("rate-console/", views.dashboard, name="dashboard"),

    # API - Current Rates
    path("api/rates/", views.get_rates_api, name="get_rates_api"),

    # Draft Operations
    path("<slug:slug>/draft/", views.save_draft, name="save_draft"),

    # Publish Operations
    path("<slug:slug>/publish/", views.publish, name="publish"),

    # Save Published Rates
    path("<slug:slug>/save-published/", views.save_published, name="save_published"),

    # Upload Published CSV
    path("<slug:slug>/upload-published/", views.upload_published, name="upload_published"),

    # Rate History
    path("api/<slug:slug>/history/", views.get_history_list, name="get_history_list"),
    path("api/history/<int:history_id>/", views.get_history_detail, name="get_history_detail"),

    # Rate Comparison Dashboards
    path("comparison/retail/", views.retail_comparison, name="retail_comparison"),
    path("comparison/bulk/", views.bulk_comparison, name="bulk_comparison"),
    path("comparison/nre-nro/", views.nre_nro_comparison, name="nre_nro_comparison"),

    # Retail Comparison APIs & Downloads
    path("api/comparison/retail/status/", views.retail_comparison_status, name="retail_comparison_status"),
    path("comparison/retail/download/", views.download_retail_comparison, name="download_retail_comparison"),

    # Bulk Comparison Downloads
    path("comparison/bulk/download/", views.download_bulk_comparison, name="download_bulk_comparison"),

    # NRE/NRO Comparison Downloads
    path("comparison/nre-nro/download/", views.download_nre_nro_comparison, name="download_nre_nro_comparison"),

    # Rate Refresh Operations
    path("refresh-all-rates/", views.refresh_all_rates, name="refresh_all_rates"),
]

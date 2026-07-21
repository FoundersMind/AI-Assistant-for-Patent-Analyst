from django.urls import path

from . import views


urlpatterns = [
    path("", views.index, name="index"),
    path("robots.txt", views.robots_txt, name="robots_txt"),
    path("api/cases", views.api_cases, name="api_cases"),
    path("api/cases/create", views.api_cases_create, name="api_cases_create"),
    path("api/cases/<int:case_id>/update", views.api_case_update, name="api_case_update"),
    path("api/cases/<int:case_id>/delete", views.api_case_delete, name="api_case_delete"),
    path("api/claim-charts/upload", views.api_claim_charts_upload, name="api_claim_charts_upload"),
    path("api/product-docs/upload", views.api_product_docs_upload, name="api_product_docs_upload"),
    path("api/product-docs/from-url", views.api_product_docs_from_url, name="api_product_docs_from_url"),
    path("api/claim-charts/<int:chart_id>/delete", views.api_claim_chart_delete, name="api_claim_chart_delete"),
    path("api/claim-charts/<int:chart_id>/update", views.api_claim_chart_update, name="api_claim_chart_update"),
    path("api/product-docs/<int:doc_id>/delete", views.api_product_doc_delete, name="api_product_doc_delete"),
    path("api/product-docs/<int:doc_id>/update", views.api_product_doc_update, name="api_product_doc_update"),
    path("api/claim-charts/<int:chart_id>/rows/update", views.api_claim_chart_row_update, name="api_claim_chart_row_update"),
    path("api/claim-charts/<int:chart_id>/rows/delete", views.api_claim_chart_row_delete, name="api_claim_chart_row_delete"),
    path("api/claim-charts/<int:chart_id>/rows/add-empty", views.api_claim_chart_row_add_empty, name="api_claim_chart_row_add_empty"),
    path("api/claim-charts/<int:chart_id>/chat/clear", views.api_claim_chart_chat_clear, name="api_claim_chart_chat_clear"),
    path("api/claim-charts/<int:chart_id>/history/clear", views.api_claim_chart_history_clear, name="api_claim_chart_history_clear"),
    path("api/claim-charts/<int:chart_id>", views.api_claim_chart_detail, name="api_claim_chart_detail"),
    path("api/claim-charts/<int:chart_id>/chat", views.api_claim_chart_chat, name="api_claim_chart_chat"),
    path(
        "api/claim-charts/<int:chart_id>/suggestions/apply",
        views.api_claim_chart_apply_suggestion,
        name="api_claim_chart_apply_suggestion",
    ),
    path("api/claim-charts/<int:chart_id>/undo", views.api_claim_chart_undo, name="api_claim_chart_undo"),
    path("api/claim-charts/<int:chart_id>/redo", views.api_claim_chart_redo, name="api_claim_chart_redo"),
    path("api/claim-charts/<int:chart_id>/history", views.api_claim_chart_history, name="api_claim_chart_history"),
    path(
        "api/claim-charts/<int:chart_id>/suggestions/apply-bulk",
        views.api_claim_chart_apply_suggestions_bulk,
        name="api_claim_chart_apply_suggestions_bulk",
    ),
    path(
        "api/claim-charts/<int:chart_id>/reassess-strengths",
        views.api_claim_chart_reassess_strengths,
        name="api_claim_chart_reassess_strengths",
    ),
    path("api/product-docs/<int:doc_id>", views.api_product_doc_detail, name="api_product_doc_detail"),
    path("api/claim-charts/<int:chart_id>/export.docx", views.api_claim_chart_export_docx, name="api_claim_chart_export_docx"),
    path("api/claim-charts/<int:chart_id>/export.csv", views.api_claim_chart_export_csv, name="api_claim_chart_export_csv"),
    path("api/claim-charts/<int:chart_id>/export.xlsx", views.api_claim_chart_export_xlsx, name="api_claim_chart_export_xlsx"),
]


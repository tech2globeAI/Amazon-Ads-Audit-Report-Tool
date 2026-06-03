"""Vercel default entrypoint — re-exports the Flask app from web_app."""

from web_app import app  # noqa: F401

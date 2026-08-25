from __future__ import annotations


class DocumentSystemError(Exception):
    """Base class for expected operator-facing failures."""


class ManifestError(DocumentSystemError):
    """A manifest cannot be loaded or validated."""


class ResolutionError(DocumentSystemError):
    """A declared selector, path, component, or profile cannot be resolved."""


class OwnershipConflict(DocumentSystemError):
    """A build would overwrite content whose canonical owner is ambiguous."""


class PackageError(DocumentSystemError):
    """A Word package is malformed or violates a required package invariant."""


class ReleaseGateError(DocumentSystemError):
    """A document is buildable but not eligible for controlled release."""


class IntegrityError(DocumentSystemError):
    """A build, current mirror, or artifact pair cannot be trusted."""


class RevisionError(DocumentSystemError):
    """A transactional canonical-source revision could not be completed safely."""


class PublishError(DocumentSystemError):
    """A verified artifact set could not be published atomically."""


class OperationPartialError(DocumentSystemError):
    """Canonical/build work succeeded, but a later optional step failed."""

    def __init__(self, message: str, report: dict):
        super().__init__(message)
        self.report = report

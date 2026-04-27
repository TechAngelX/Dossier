using Dossier.Models;

namespace Dossier.Services;

public interface IPdfRenameService
{
    List<PdfRenamePreviewItem> PreviewRenames(string folderPath, IEnumerable<StudentRecord> students);

    // Renames files in-place within the same folder (used for auto-rename during download)
    (int success, int errors, List<(string newPath, string originalPath)> completed) RenameInPlace(string folderPath, IReadOnlyList<PdfRenamePreviewItem> items);

    // Renames files and moves them into a named batch subfolder
    (int success, int errors, List<(string destPath, string originalPath)> completed) RenameIntoBatchFolder(string folderPath, IReadOnlyList<PdfRenamePreviewItem> items, string batchFolderName);

    (int success, int skipped, int errors) AppendRanking(string folderPath, IEnumerable<StudentRecord> students);
    string GenerateNewFilename(StudentRecord student);
}

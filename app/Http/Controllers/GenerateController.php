<?php

namespace App\Http\Controllers;

use PDF;
use Illuminate\Support\Str;
use PhpOffice\PhpWord\Settings;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\TemplateProcessor;

class GenerateController extends Controller
{
    /**
     * Download sebagai .docx
     */
    public function download()
    {
        $name = 'Yuri Dimas Satrio';
        $email = 'yuridimas00@gmail.com';

        $template = new TemplateProcessor(storage_path('Sertifikat.docx'));

        $template->setValue('nama', $name);
        $template->setValue('email', $email);

        $filename = 'Sertifikas_' . Str::slug($name) . '.docx';

        header('Content-Type: application/octet-stream');
        header("Content-Disposition: attachment; filename=$filename");

        $template->saveAs('php://output');
    }

    /**
     * Print sebagai .pdf
     */
    public function print()
    {
        Settings::setPdfRendererPath('../vendor/dompdf/dompdf');
        Settings::setPdfRendererName('DomPDF');

        $name = 'Yuri Dimas Satrio';
        $email = 'yuridimas00@gmail.com';

        $template = new TemplateProcessor(storage_path('Sertifikat.docx'));

        $template->setValue('nama', $name);
        $template->setValue('email', $email);

        $filename = 'Sertifikat_' . Str::slug($name);
        $path = storage_path($filename . '.docx');

        $template->saveAs($path);

        $temp = IOFactory::load($path);
        $xmlWriter = IOFactory::createWriter($temp, 'PDF');

        $xmlWriter->save(storage_path($filename . '.pdf'), TRUE);

        return response()
            ->download(storage_path($filename . '.pdf'));
    }
}

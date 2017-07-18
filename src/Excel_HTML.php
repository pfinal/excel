<?php

namespace PFinal\Excel;

/**
 * PHPExcel_Writer_HTML
 *
 * @category   PHPExcel
 * @package    PHPExcel_Writer_HTML
 */
class Excel_HTML extends \PHPExcel_Writer_HTML
{
    /**
     * Create a new PHPExcel_Writer_HTML
     *
     * @param    \PHPExcel $phpExcel PHPExcel object
     */
    public function __construct(\PHPExcel $phpExcel)
    {
        parent::__construct($phpExcel);
    }

    public function ToHtml()
    {
        $html = '';
        $this->_phpExcel->garbageCollect();
        $this->buildCSS(!$this->getUseInlineCss());
        $html .= $this->generateStyles(true);
        $html .= $this->generateSheetData();
        return $html;
    }
}

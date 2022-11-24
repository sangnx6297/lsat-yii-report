<?php

namespace lcssoft\report\handler;

use lcssoft\report\helpers\Logger;
use lcssoft\report\helpers\Utilities;
use Mpdf\Tag\P;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use phpseclib3\Crypt\Hash;
use yii\base\Model;
use yii\db\ActiveQuery;
use yii\db\Exception;
use yii\helpers\Inflector;

class ReportHandler extends \yii\base\BaseObject
{
    public $objectClass;
    public $parameters;

    /**
     * @return bool
     * @throws Exception
     * @throws \yii\base\InvalidConfigException
     */
    public function exec()
    {

        $object = \Yii::createObject($this->objectClass);

        if (!($object instanceof ReportTemplateInterface)) {
            throw new Exception("Object not instance of ReportTemplateInterface");
        }
        $query = $object->getQuery($this->parameters);
        Logger::log($query->createCommand()->rawSql);
        $columns = $object->columns();
        return $this->handler($object, $query, $this->initColumns($columns, $query));
    }

    /**
     * @param $object Model | ReportTemplateInterface
     * @param $query ActiveQuery
     * @param $columns ReportColumnData[]
     * @return boolean
     */
    public function handler($object, $query, $columns)
    {
        $template = \yii\helpers\Url::to('@rootPath/sample/report-template.xlsx');
        if ($tempTemplate = $object->getReportFileTemplate()) {
            $template = $tempTemplate;
        }

        $objExcelTemplate = IOFactory::load($template);
        $activeSheet = $objExcelTemplate->getActiveSheet();
        $rowIndex = 1;

        $activeSheet = $this->renderSheetHeader($activeSheet, $object, $columns, $rowIndex);

        $activeSheet = $this->renderSheetBody($query, $activeSheet, $columns, $rowIndex);

        $objWriter = IOFactory::createWriter($activeSheet->getParent(), 'Xlsx');
        $basePath = \yii\helpers\Url::to('@rootPath/static/download');
        $basePath = Utilities::createDirectory([$basePath, date('Y'), date('m'), get_class($object)]);


        $filePath = $basePath . DIRECTORY_SEPARATOR . $this->getFileName($object) . date('Y_m_d_h_i_s') . '.xlsx';
        $objWriter->save($filePath);
        return $filePath;
    }

    /**
     * @param $columns
     * @param $query ActiveQuery
     * @return ReportColumnData[]
     */
    public function initColumns($columns, $query)
    {
        $columnData = [];

        if (empty($columns)) {
            throw new \yii\base\Exception("Columns must be declare.");
        }

        $columnData[] = $this->getIndexColumn();

        foreach ($columns as $column) {

            $columnConfig = [];

            if (is_array($column)) {
                $columnConfig = array_merge($columnConfig, $column);
            } else {
                $columnConfig = [
                    'attribute' => $column,
                ];
            }

            if (!isset($columnConfig['class'])) {
                $columnConfig['class'] = ReportColumnData::class;
            }

            $columnData[] =  \Yii::createObject($columnConfig);
        }

        return $columnData;
    }

    public function getIndexColumn()
    {
        $config = [
            'class' => ReportColumnData::class,
            'label' => \Yii::t('backend', 'Index'),
            'value' => function ($model, $cell, $index, $rowIndex) {
                return $index + 1;
            }
        ];

        return \Yii::createObject($config);
    }

    /**
     * @param $activeSheet Worksheet
     * @param $object Model
     * @param $columns ReportColumnData[]
     * @return Worksheet
     */
    public function renderSheetHeader($activeSheet, $object, $columns, &$rowIndex)
    {
        $rowIndex = $rowIndex++;

        echo PHP_EOL;
        echo PHP_EOL;


        foreach ($columns as $idx => $column) {
            $colStr = Coordinate::stringFromColumnIndex(($idx+1));
            $cellKey = "{$colStr}{$rowIndex}";
            if($column->headerOptions){
                $activeSheet->getCell($cellKey)->getStyle()->applyFromArray($column->headerOptions, false);
                $activeSheet->getColumnDimension($colStr)->setAutoSize($column->autoResize);
            }

            $activeSheet->getCell($cellKey)->setValue($column->getLabel($object));
        }

        return $activeSheet;
    }

    /**
     * @param $query ActiveQuery
     * @param $activeSheet Worksheet
     * @param $columns ReportColumnData []
     * @param $rowIndex integer
     * @return Worksheet
     * @throws \yii\base\InvalidConfigException
     */
    public function renderSheetBody($query, $activeSheet, $columns, &$rowIndex)
    {

        $statement = $query->createCommand()->query();
        $rowIndex++;
        $index = 0;
        while ($objectData = $statement->read()) {
            $objectData['class'] = $this->objectClass;
            $object = \Yii::createObject($objectData);

            $colIndex = 1;
//            $colStr = Coordinate::stringFromColumnIndex($colIndex++);
//            $activeSheet->getCell("{$colStr}{$rowIndex}")->setValue(($index + 1));

            foreach ($columns as $idx => $column) {

                $colStr = Coordinate::stringFromColumnIndex($colIndex++);
                $cell = $column->renderAttribute($activeSheet->getCell("{$colStr}{$rowIndex}"), $object, $index, $rowIndex);

                $args = [
                    'model' => $object,
                    'cell' => $cell,
                    'index' => $index,
                    'rowIndex' => $rowIndex,
                ];


                $cellKey = "{$colStr}{$rowIndex}";

                if ($column->contentOptions) {
                    $activeSheet->getCell($cellKey)->getStyle()->applyFromArray($column->contentOptions);
                }
                $activeSheet->getCell($cellKey)->setValue($column->formattedValue($cell->getValue(), $args));

                if($idx == 0){
                    $activeSheet->getColumnDimension($colStr)->setAutoSize($column->autoResize);
                    $activeSheet->getColumnDimension($colStr)->setWidth($column->getAttributeValue('visible',$args));
                    $activeSheet->getColumnDimension($colStr)->setVisible($column->getAttributeValue('width',$args));
                }

                echo $cell->getValue() . "\t|\t";
            }
            echo PHP_EOL;

            $index++;
            $rowIndex++;
        }
        return $activeSheet;
    }


    /**
     * @param $model Model
     * @param $columns ReportColumnData[]
     * @return mixed;
     */
    public function getHeaderLabels($model, $columns)
    {
        $headerLabels = [];
        foreach ($columns as $column) {
            $headerLabels[] = $column->getLabel($model);
        }
        return $headerLabels;
    }

    /**
     * @param $object Model | ReportTemplateInterface
     * @return string
     */
    public function getFileName($object)
    {
        $fileName = $object->getReportFileName();
        if (empty($fileName)) {
            return md5(microtime(false));
        }
        return $fileName;
    }


}
<?php

namespace lcssoft\report\handler;

use yii\db\ActiveQuery;

interface ReportTemplateInterface
{

    /**
     * @param $parameters
     * @return ActiveQuery
     */
    public function getQuery($parameters);

    /**
     * @desc
     *  @array [
     *              'class' => ReportColumnData::class,
     *              'attribute' => '',
     *              'value' => string | closure function
     *              'label' => '',
     *              'format' => ''
     *         ]
     * @return mixed
     */
    public function columns();

    /**
     * @return string
     */
    public function getReportFileName():string;

    /**
     * @return string
     */
    public function getReportFileTemplate():string;

    /**
     * @return string
     */
    public function getReportLabel():string;

}
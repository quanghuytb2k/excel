<?php
namespace App\Http\Controllers;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\IReadFilter;
use Maatwebsite\Excel\Facades\Excel;
use Illuminate\Http\Request;
use DB;

class ExportExcelController extends Controller
{
    function index()
    {



     $customer_data = DB::table('tbl_customer')->get();
     return view('export_excel')->with('customer_data', $customer_data);
    }

    function excel()
    {
     $customer_data = DB::table('tbl_customer')->get()->toArray();
     $customer_array[] = array('Customer Name', 'Address', 'City', 'Postal Code', 'Country');
     foreach($customer_data as $customer)
     {
      $customer_array[] = array(
       'Customer Name'  => $customer->CustomerName,
       'Address'   => $customer->Address,
       'City'    => $customer->City,
       'Postal Code'  => $customer->PostalCode,
       'Country'   => $customer->Country
      );
     }
     echo"<pre>";
     print_r($customer_array);
     Excel::create('Customer Data', function($excel) use ($customer_array){
      $excel->setTitle('Customer Data');
      $excel->sheet('Customer Data', function($sheet) use ($customer_array){
       $sheet->fromArray($customer_array, null, 'A1', false, false);
      });
     })->download('xlsx');
    }

    public function export()
    {
        $salarys = DB::table('tbl_customer')->get()->toArray();
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('report.xls');
        $worksheet = $spreadsheet->getActiveSheet();
        $i=20;
        if($salarys != null){
            foreach ($salarys as $key => $value) {
                $worksheet->getCell('B'.$i)->setValue($value->CustomerName);
                $worksheet->getCell('C'.$i)->setValue($value->Address);
                // $worksheet->getCell('C'.$i)->setValue($value->City);
                // $worksheet->getCell('D'.$i)->setValue($value->PostalCode);
                // $worksheet->getCell('E'.$i)->setValue($value->Country);
                $i++;
                $worksheet->insertNewRowBefore($i); //thêm hàng
            }
        }
        $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xls');
        $writer->save('report.xls');
        return redirect('report.xls');
        // return Excel::download(new SalaryExport, 'salaries.xlsx');
    }

}

?>

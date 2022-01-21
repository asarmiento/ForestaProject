<?php


namespace App\Http\Controllers;


use App\Imports\ImportNewExcelFarm;
use App\Models\Family;
use App\Models\ForestDataBase;
use App\Models\ScientificName;
use Illuminate\Http\Request;
use Maatwebsite\Excel\Facades\Excel;
use PhpOffice\PhpWord\SimpleType\TblWidth;

class FarmController extends Controller
{

    public function importStore(Request $request)
    {
        $file=$request->file();
        $import=new ImportNewExcelFarm;

        Excel::import($import,$file['data_farm']);
        return redirect()->back();
    }


    public function listsFamily()
    {
        $list=Family::all();
        return view('Families.lists',compact('list'));
    }
}

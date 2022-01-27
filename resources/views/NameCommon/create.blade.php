@extends('layouts.app')

@section('content')
    <div class="container">
        <div class="row justify-content-center">
            <div class="col-md-8">
                <div class="card">
                    <div class="card-header bg-dark">{{ __('Registro de Nombres comunes') }}</div>

                    <div class=" card-body">
        <form method="post" action="{{route('store-family')}}">
           <div class="row">
               <div class="col-md-6">
                   <label>Nombre de Cientifico</label>
                   <select name="" class="form-control">
                       @foreach($scientifics AS $scientific)
                           <option value="{{$scientific->id}}">{{$scientific->name}}</option>
                       @endforeach
                   </select>
               </div>
               <div class="col-md-6">
                   <label>Nombre Común</label>
                   <input type="text" id="name" name="name" class="form-control" />
               </div>
               <div class="col-md-12 text-center m-4">
                   <button type="submit" class="btn btn-success">Guardar</button>
               </div>

           </div>
        </form>
    </div>
    </div>
@endsection
@section('script')
    <script src="{{ mix('js/app.js') }}" defer></script>
@endsection

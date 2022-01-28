@extends('layouts.app')

@section('content')
    <div class="container">
        <div class="row justify-content-center">
            <div class="col-md-8">
                <div class="card">
                    <div class="card-header bg-dark text-white" >{{ __('Registro de Nombres Científicos') }}</div>

                    <div class=" card-body">
        <form method="post" action="{{route('update-scientific',$scientific->id)}}">
            @method('put')
           <div class="row">{{csrf_field()}}
               <div class="col-md-4">
                   <label>Nombre de Familia</label>
                   <select name="family_id" class="form-control">

                       <option value="{{$scientific->family->id}}">{{$scientific->family->name}}</option>
                       @foreach($family AS $scientific)
                           <option value="{{$scientific->id}}">{{$scientific->name}}</option>
                       @endforeach
                   </select>
               </div>
               <div class="col-md-4">
                   <label>Nombre Científico</label>
                   <input type="text" id="name" name="name" value="{{$scientific->name}}" class="form-control" />
               </div>
               <div class="col-md-4">
                   <label>Tipo de Especie</label>
                   <select name="commercial" class="form-control">
                       @if($scientific->commercial == 1)
                           <option  value="1" >{{$scientific->commercial}} Comercial</option>
                       @else
                           <option  value="0" >{{$scientific->commercial}} No Comercial</option>
                       @endif
                           <option value="1">Comercial</option>
                           <option value="0">No Comercial</option>
                   </select>
               </div>
               <div class="col-md-12 text-center m-4">
                   <button type="submit" class="btn btn-success">Actualizar</button>
               </div>

           </div>
        </form>
    </div>
    </div>
@endsection
@section('script')
    <script src="{{ mix('js/app.js') }}" defer></script>
@endsection

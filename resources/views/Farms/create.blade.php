@extends('layouts.app')

@section('content')
    <div class="container">
        <div class="row justify-content-center">
            <div class="col-md-8">
                <div class="card">
                    <div class="card-header bg-dark text-white">{{ __('Registro de Fincas') }}</div>

                    <div class=" card-body">
        <form method="post" action="{{route('store-farm')}}">
           <div class="row">{{csrf_field()}}
               <div class="col-md-6">
                   <label>ID Predio</label>
                   <input type="text" id="id_predio" name="id_predio" class="form-control" />
               </div>
               <div class="col-md-6">
                   <label>Oficinas SINAC</label>
                   <input type="text" id="name" name="office_sinac" class="form-control" />
               </div>
               <div class="col-md-6">
                   <label>Cantidad Vano</label>
                   <input type="text" id="name" name="count_vano" class="form-control" />
               </div>
               <div class="col-md-6">
                   <label>Detalles Vanos</label>
                   <input type="text" id="name" name="detail_vano" class="form-control" />
               </div>
               <div class="col-md-6">
                   <label>Cédula de Propietario</label>
                   <input type="text" id="name" name="card" class="form-control" />
               </div>
               <div class="col-md-6">
                   <label>Propietario</label>
                   <input type="text" id="name" name="owner" class="form-control" />
               </div>
               <div class="col-md-6">
                   <label>Folio Real</label>
                   <input type="text" id="name" name="folio_real" class="form-control" />
               </div>
               <div class="col-md-6">
                   <label>Plano</label>
                   <input type="text" id="name" name="plane" class="form-control" />
               </div>
               <div class="col-md-6">
                   <label>Cita Contrato Servidumbre</label>
                   <input type="text" id="name" name="appointment_contract" class="form-control" />
               </div>
               <div class="col-md-6">
                   <label>Longitud KM</label>
                   <input type="text" id="name" name="logitud_km" class="form-control" />
               </div>
               <div class="col-md-6">
                   <label>Predio Anterior</label>
                   <input type="text" id="name" name="predio_before" class="form-control" />
               </div>
               <div class="col-md-6">
                   <label>Predio Posterior</label>
                   <input type="text" id="name" name="predio_after" class="form-control" />
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

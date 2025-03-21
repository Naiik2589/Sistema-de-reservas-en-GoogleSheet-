function validarReservas() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var disponibilidad = [];
  var conflictos = [];

  Logger.log("Inicio de la función validarReservas");

  for (var i = 1; i < data.length; i++) {
    var fila = data[i];
    var fechaReserva = fila[3];
    var horaInicio = fila[4];
    var horaFin = fila[5];
    var salaReuniones = fila[6];
    var responsable = fila[1];
    var direccionCorreo = fila[9];
    var notificado = fila[10];

    var fechaReservaFormateada = Utilities.formatDate(fechaReserva, "GMT-0500", "dd/MM/yyyy");
    Logger.log("Fecha formateada: " + fechaReservaFormateada);

    // Horario laboral
    var horarioLaboral = [
      { inicio: "07:00", fin: "17:00" }
    ];

    var disponible = true;

    for (var j = 0; j < disponibilidad.length; j++) {
      if (
        disponibilidad[j].fecha === fechaReservaFormateada &&
        disponibilidad[j].sala === salaReuniones
      ) {
        if (
          (horaInicio >= disponibilidad[j].horaInicio && horaInicio < disponibilidad[j].horaFin) ||
          (horaFin > disponibilidad[j].horaInicio && horaFin <= disponibilidad[j].horaFin)
        ) {
          disponible = false;
          break;
        }
      }
    }

    if (disponible) {
      disponibilidad.push({
        fecha: fechaReservaFormateada,
        sala: salaReuniones,
        horaInicio: horaInicio,
        horaFin: horaFin
      });

      Logger.log("Reserva aprobada: " + horaInicio + " - " + horaFin);
      sheet.getRange(i + 1, 12).setValue("Sí");

      if (notificado !== "Si") {
        MailApp.sendEmail(direccionCorreo, "Reserva de sala confirmada", "Hola " + responsable + ",\n\nTu reserva para el " + fechaReservaFormateada + ", de " + horaInicio + " a " + horaFin + " en la " + salaReuniones + " ha sido confirmada.\n\nFoscal \nInpirados por la vida");
        sheet.getRange(i + 1, 11).setValue("Si");
      }
    } else {
      Logger.log("Conflicto detectado para: " + horaInicio + " - " + horaFin);
      sheet.getRange(i + 1, 12).setValue("No");

      // Calcular horarios disponibles
      var horariosDisponibles = calcularHorariosDisponibles(horarioLaboral, disponibilidad, fechaReservaFormateada, salaReuniones);
      Logger.log("Horarios disponibles: " + JSON.stringify(horariosDisponibles));

      if (notificado !== "Si") {
        var mensaje = "Hola " + responsable + ",\n\nTu reserva no fue aprobada debido a que ya hay una reserva para esa fecha y hora.";
        mensaje += "\n\nHorarios disponibles para ese día:";
        horariosDisponibles.forEach(function (horario) {
          mensaje += "\n" + horario.inicio + " - " + horario.fin;
        });

        MailApp.sendEmail(direccionCorreo, "Conflicto de reserva", mensaje);
        sheet.getRange(i + 1, 11).setValue("Si");
      }
    }
  }

  Logger.log("Fin de la función validarReservas");
}

function calcularHorariosDisponibles(horarioLaboral, reservas, fecha, sala) {
  Logger.log("Calculando horarios disponibles...");

  var ocupados = reservas.filter(r => r.fecha === fecha && r.sala === sala);
  ocupados.sort((a, b) => a.horaInicio.localeCompare(b.horaInicio));

  var disponibles = [];
  var inicioDisponible = horarioLaboral[0].inicio;

  ocupados.forEach(function (reserva) {
    if (inicioDisponible < reserva.horaInicio) {
      disponibles.push({ inicio: inicioDisponible, fin: reserva.horaInicio });
    }
    inicioDisponible = reserva.horaFin;
  });

  if (inicioDisponible < horarioLaboral[0].fin) {
    disponibles.push({ inicio: inicioDisponible, fin: horarioLaboral[0].fin });
  }

  Logger.log("Horarios disponibles calculados: " + JSON.stringify(disponibles));
  return disponibles;
}

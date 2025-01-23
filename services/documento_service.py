from models.document_request import DocumentRequest
from docx import Document
from docxtpl import DocxTemplate

class DocumentoService:
    
    @staticmethod
    def process_request(request: DocumentRequest):
        #Carga del documento
        document = DocxTemplate('lib/modelo.docx')
        condicion = DocumentoService.validacion_condicion(request, document)
        return condicion
    
    @staticmethod
    def validacion_condicion(request: DocumentRequest, document: Document):
        if request.condicion not in ['contado', 'financiado', 'fraccionado']:
            raise ValueError("Condición no válida. Usa 'contado', 'financiado' o 'fraccionado'.")
        DocumentoService.tipo_condicion(request.condicion, request, document)
    
    @staticmethod
    def tipo_condicion(condicion: str, request: DocumentRequest, document: Document):
        if condicion == 'contado':
            return DocumentoService.counted_type(request, document)
        elif condicion == 'financiado':
            return DocumentoService.financed_type(request, document)
        elif condicion == 'fraccionado':
            return DocumentoService.fractionated_type(request, document)
        else:
            raise ValueError("Condición desconocida.")
    
    @staticmethod
    def counted_type(request: DocumentRequest, document: Document):
        
        # Aquí se define el diccionario con los valores que se reemplazarán en el documento
        valores = {
            # Texto estático
            'texto_4': 'La entrega de la posesión de el/los lote(s) se realizará en el mes de diciembre de 202x.',
            'texto_5': 'La entrega de la posesión de las áreas y servicios comunes del Condominio se realizará en el mes de diciembre de 202x.',
            'texto_7': '',
            'texto_8': '',
            'texto_9': '',
            'texto_10': 'Según el cronograma de pago indicado en el Numeral 10 del Anexo 5: Hoja Resumen',

            # Fecha
            'day': request.day or '',
            'month': request.month or '',
            'year': request.year or '',

            # Datos del cliente 1
            'name_1': request.name_1 or '',
            'dni_1': request.dni_1 or '',
            'ocupation_1': request.ocupation_1 or '',
            'marital_status_1': request.marital_status_1 or '',
            'address_1': request.address_1 or '',
            'mail_1': request.mail_1 or '',
            'phone_1': request.phone_1 or '',

            # Datos del cliente 2
            'name_2': request.name_2 or '',
            'dni_2': request.dni_2 or '',
            'ocupation_2': request.ocupation_2 or '',
            'marital_status_2': request.marital_status_2 or '',
            'address_2': request.address_2 or '',
            'mail_2': request.mail_2 or '',
            'phone_2': request.phone_2 or '',

            # Datos del lote
            'number_batch': request.number_batch or '',
            'approximate_area': request.approximate_area or '',
        }
        
        document.render(valores)
        
        DocumentoService.reemplazar_marcadores(document, valores)
        DocumentoService.eliminar_parrafos_innecesarios(document)
        DocumentoService.eliminar_desde_marcador(document, '${eliminar}')
        
        document.save('documento_contado_final.docx')
    
    @staticmethod
    def financed_type(request: DocumentRequest, document: Document):
        
        valores = {

        #financed according to the table
        # 'texto_1':'Anexo Nº 5: Hoja Resumen',
        # 'texto_2':'El Comprador declara conocer que las indicadas son cuentas recaudadoras razón por la que ante el incumplimiento de pago en la fecha correspondiente incurrirá en mora automática sin necesidad de intimación previa; en consecuencia, se devengará un interés compensatorio diario de US$ 1.00 (Un y 00/100 dólares americanos), y un interés moratorio diario igual, ambos respecto del importe de la cuota adeudada, los cuales se cobrarán conjuntamente con la cuota pendiente de pago. El Comprador reconoce que los pagos deben efectuarse, obligatoriamente, a través de dicha cuenta recaudadora, considerándose esta como una obligación contractual <br> Sin perjuicio de ello, El Comprador declara conocer que, supletoriamente al sistema de recaudación mencionado, podrá realizar el pago de las cuotas mediante el acceso a un enlace de pago generado por la Vendedora y/o sistema de recaudación propuesto por la Vendedora; el mismo que también generará una mora automática compuesta por un interés compensatorio diario y un interés moratorio diario del mismo valor señalado en el párrafo anterior, siempre que incumple con el pago de la cuota en la fecha correspondiente. Las partes declaran que esta forma de pago también se considerara una obligación contractual y generará los efectos cancelatorios correspondientes. <br> Finalmente, el Comprador deberá informar y enviar a la Vendedora, los sustentos de pagos respectivos.',
        # 'texto_3':'Adicionalmente, las partes dejan constancia que, al amparo de lo dispuesto por el artículo 1583 del Código Civil, la Vendedora se reserva la propiedad de el/los lote(s) hasta la cancelación total del Precio de Venta.',
        'texto_4':'La entrega de la posesión de el/los lote(s) se realizará en el mes de diciembre de 202x.',
        'texto_5':'La entrega de la posesión de las áreas y servicios comunes del Condominio se realizará en el mes de diciembre de 202x.',
        # 'texto_6':'La Vendedora podrá reportar a las centrales de riesgo a El Comprador en caso de incumplimiento en el pago de sus cuotas.',
        'texto_7':'(a) dos o más armadas alternas o consecutivas (cuotas) del Precio de Venta adeudado bajo el presente Contrato señaladas en el Cronograma de Pagos indicado en el Numeral 10 del Anexo N.° 5: Hoja Resumen; y/o (b)',
        'texto_8':'Así, en caso el Comprador mantenga algún reclamo que esté siendo materia de controversia no podrá suspender el pago de las cuotas del financiamiento que mantenga pendientes en atención al lote adquirido ni podrá suspender las demás obligaciones que haya contraído, salvo que cuente con una orden judicial o arbitral que así lo determine.',
        'texto_9':'El saldo de US$ --- (--- con 00/100 dólares americanos), que será cancelado',
        'texto_10': 'según el cronograma de pago indicado en el Numeral 10 del Anexo 5: Hoja Resumen', 
        'texto_11': '',
        'texto_12': '',
        
        #day and month
        'day': request.day or '',
        'month': request.month or '',
        'year': request.year or '',

        # Datos del cliente 1
        'name_1': request.name_1 or '',
        'dni_1': request.dni_1 or '',
        'ocupation_1': request.ocupation_1 or '',
        'marital_status_1': request.marital_status_1 or '',
        'address_1': request.address_1 or '',
        'mail_1': request.mail_1 or '',
        'phone_1': request.phone_1 or '',

        # Datos del cliente 2
        'name_2': request.name_2 or '',
        'dni_2': request.dni_2 or '',
        'ocupation_2': request.ocupation_2 or '',
        'marital_status_2': request.marital_status_2 or '',
        'address_2': request.address_2 or '',
        'mail_2': request.mail_2 or '',
        'phone_2': request.phone_2 or '',

        # Datos del lote
        'number_batch': request.number_batch or '',
        'approximate_area': request.approximate_area or '',
        }
        
        document.render(valores)
        
        DocumentoService.insertar_texto_estatico(document)
        DocumentoService.reemplazar_marcadores(document, valores)
        DocumentoService.dejar_el_marcador(document)
        
        document.save('documento_financiado_final.docx')
    
    @staticmethod
    def fractionated_type(request: DocumentRequest, document: Document):
        valores = {

        #financed according to the table
        # 'texto_2':'El Comprador declara conocer que las indicadas son cuentas recaudadoras razón por la que ante el incumplimiento de pago en la fecha correspondiente incurrirá en mora automática sin necesidad de intimación previa; en consecuencia, se devengará un interés compensatorio diario de US$ 1.00 (Un y 00/100 dólares americanos), y un interés moratorio diario igual, ambos respecto del importe de la cuota adeudada, los cuales se cobrarán conjuntamente con la cuota pendiente de pago. El Comprador reconoce que los pagos deben efectuarse, obligatoriamente, a través de dicha cuenta recaudadora, considerándose esta como una obligación contractual <br> Sin perjuicio de ello, El Comprador declara conocer que, supletoriamente al sistema de recaudación mencionado, podrá realizar el pago de las cuotas mediante el acceso a un enlace de pago generado por la Vendedora y/o sistema de recaudación propuesto por la Vendedora; el mismo que también generará una mora automática compuesta por un interés compensatorio diario y un interés moratorio diario del mismo valor señalado en el párrafo anterior, siempre que incumple con el pago de la cuota en la fecha correspondiente. Las partes declaran que esta forma de pago también se considerara una obligación contractual y generará los efectos cancelatorios correspondientes. <br> Finalmente, el Comprador deberá informar y enviar a la Vendedora, los sustentos de pagos respectivos.',
        # 'texto_3':'Adicionalmente, las partes dejan constancia que, al amparo de lo dispuesto por el artículo 1583 del Código Civil, la Vendedora se reserva la propiedad de el/los lote(s) hasta la cancelación total del Precio de Venta.',
        'texto_4':'La entrega de la posesión de el/los lote(s) se realizará en el mes de diciembre de 202x.',
        'texto_5':'La entrega de la posesión de las áreas y servicios comunes del Condominio se realizará en el mes de diciembre de 202x.',
        # 'texto_6':'La Vendedora podrá reportar a las centrales de riesgo a El Comprador en caso de incumplimiento en el pago de sus cuotas.',
        'texto_7':'(a) dos o más armadas alternas o consecutivas (cuotas) del Precio de Venta adeudado bajo el presente Contrato señaladas en el Cronograma de Pagos indicado en el Numeral 10 del Anexo N.° 5: Hoja Resumen; y/o (b)',
        'texto_8':'Así, en caso el Comprador mantenga algún reclamo que esté siendo materia de controversia no podrá suspender el pago de las cuotas del financiamiento que mantenga pendientes en atención al lote adquirido ni podrá suspender las demás obligaciones que haya contraído, salvo que cuente con una orden judicial o arbitral que así lo determine.',
        'texto_9':'El saldo de US$ --- (--- con 00/100 dólares americanos), que será cancelado',
        'texto_10':'',
        'texto_11': 'según lo siguiente:',
        'texto_12': '(i)	La suma de US$ --- (--- con 00/100 dólares americanos), a más tardar el – de – de 202-. \n (ii)	La suma de US$ --- (--- con 00/100 dólares americanos), a más tardar el – de – de 202-. \n (iii)	(….)',
        
        #day and month
        'day': request.day or '',
        'month': request.month or '',
        'year': request.year or '',

        # Datos del cliente 1
        'name_1': request.name_1 or '',
        'dni_1': request.dni_1 or '',
        'ocupation_1': request.ocupation_1 or '',
        'marital_status_1': request.marital_status_1 or '',
        'address_1': request.address_1 or '',
        'mail_1': request.mail_1 or '',
        'phone_1': request.phone_1 or '',

        # Datos del cliente 2
        'name_2': request.name_2 or '',
        'dni_2': request.dni_2 or '',
        'ocupation_2': request.ocupation_2 or '',
        'marital_status_2': request.marital_status_2 or '',
        'address_2': request.address_2 or '',
        'mail_2': request.mail_2 or '',
        'phone_2': request.phone_2 or '',

        # Datos del lote
        'number_batch': request.number_batch or '',
        'approximate_area': request.approximate_area or '',
        }
        
        document.render(valores)
        
        DocumentoService.insertar_texto_estatico_fractionated(document)
        DocumentoService.reemplazar_marcadores(document, valores)
        DocumentoService.eliminar_parrafos_innecesarios_fractioned(document)
        DocumentoService.eliminar_desde_marcador(document, '${eliminar}')    
        
        document.save('documento_fraccionado_final.docx')
    
    @staticmethod
    def insertar_texto_estatico(document):
        valores= {
        '${texto_1}':'Anexo Nº 5: Hoja Resumen',
        '${texto_2}':'El Comprador declara conocer que las indicadas son cuentas recaudadoras razón por la que ante el incumplimiento de pago en la fecha correspondiente incurrirá en mora automática sin necesidad de intimación previa; en consecuencia, se devengará un interés compensatorio diario de US$ 1.00 (Un y 00/100 dólares americanos), y un interés moratorio diario igual, ambos respecto del importe de la cuota adeudada, los cuales se cobrarán conjuntamente con la cuota pendiente de pago. El Comprador reconoce que los pagos deben efectuarse, obligatoriamente, a través de dicha cuenta recaudadora, considerándose esta como una obligación contractual <br> Sin perjuicio de ello, El Comprador declara conocer que, supletoriamente al sistema de recaudación mencionado, podrá realizar el pago de las cuotas mediante el acceso a un enlace de pago generado por la Vendedora y/o sistema de recaudación propuesto por la Vendedora; el mismo que también generará una mora automática compuesta por un interés compensatorio diario y un interés moratorio diario del mismo valor señalado en el párrafo anterior, siempre que incumple con el pago de la cuota en la fecha correspondiente. Las partes declaran que esta forma de pago también se considerara una obligación contractual y generará los efectos cancelatorios correspondientes. <br> Finalmente, el Comprador deberá informar y enviar a la Vendedora, los sustentos de pagos respectivos.',
        '${texto_3}':'Adicionalmente, las partes dejan constancia que, al amparo de lo dispuesto por el artículo 1583 del Código Civil, la Vendedora se reserva la propiedad de el/los lote(s) hasta la cancelación total del Precio de Venta.',
        '${texto_6}':'La Vendedora podrá reportar a las centrales de riesgo a El Comprador en caso de incumplimiento en el pago de sus cuotas.',
        }
        
        # Iterar sobre todos los párrafos del documento y reemplazar los marcadores
        for paragraph in document.paragraphs:
            for marcador, reemplazo in valores.items():
                if marcador in paragraph.text:  # Verifica si el marcador está en el párrafo
                    # Reemplaza el marcador con el texto proporcionado
                    paragraph.text = paragraph.text.replace(marcador, reemplazo)
    
    @staticmethod
    def insertar_texto_estatico_fractionated(document):
        valores_estaticos = {
        '${texto_2}':'El Comprador declara conocer que las indicadas son cuentas recaudadoras razón por la que ante el incumplimiento de pago en la fecha correspondiente incurrirá en mora automática sin necesidad de intimación previa; en consecuencia, se devengará un interés compensatorio diario de US$ 1.00 (Un y 00/100 dólares americanos), y un interés moratorio diario igual, ambos respecto del importe de la cuota adeudada, los cuales se cobrarán conjuntamente con la cuota pendiente de pago. El Comprador reconoce que los pagos deben efectuarse, obligatoriamente, a través de dicha cuenta recaudadora, considerándose esta como una obligación contractual <br> Sin perjuicio de ello, El Comprador declara conocer que, supletoriamente al sistema de recaudación mencionado, podrá realizar el pago de las cuotas mediante el acceso a un enlace de pago generado por la Vendedora y/o sistema de recaudación propuesto por la Vendedora; el mismo que también generará una mora automática compuesta por un interés compensatorio diario y un interés moratorio diario del mismo valor señalado en el párrafo anterior, siempre que incumple con el pago de la cuota en la fecha correspondiente. Las partes declaran que esta forma de pago también se considerara una obligación contractual y generará los efectos cancelatorios correspondientes. <br> Finalmente, el Comprador deberá informar y enviar a la Vendedora, los sustentos de pagos respectivos.',
        '${texto_3}':'Adicionalmente, las partes dejan constancia que, al amparo de lo dispuesto por el artículo 1583 del Código Civil, la Vendedora se reserva la propiedad de el/los lote(s) hasta la cancelación total del Precio de Venta.',
        '${texto_6}':'La Vendedora podrá reportar a las centrales de riesgo a El Comprador en caso de incumplimiento en el pago de sus cuotas.',
        }
        
        # Iterar sobre todos los párrafos del documento y reemplazar los marcadores
        for paragraph in document.paragraphs:
            for marcador, reemplazo in valores_estaticos.items():
                if marcador in paragraph.text:  # Verifica si el marcador está en el párrafo
                    # Reemplaza el marcador con el texto proporcionado
                    paragraph.text = paragraph.text.replace(marcador, reemplazo)
    
    # Función para reemplazar los marcadores con los valores correspondientes en el documento
    @staticmethod
    def reemplazar_marcadores(document, valores):
        for paragraph in document.paragraphs:
            for key, value in valores.items():
                if key in paragraph.text:
                    # Reemplazar el marcador con el valor
                    inline = paragraph.runs
                    for run in inline:
                        if key in run.text:
                            run.text = run.text.replace(key, value)
        
    # Función para eliminar todos los párrafos después de un marcador
    @staticmethod
    def eliminar_desde_marcador(document, marcador):
        eliminar = False
        for paragraph in document.paragraphs:
            if marcador in paragraph.text:
                eliminar = True  # Cuando encontramos el marcador, comenzamos a eliminar
            if eliminar:
                # Elimina el párrafo
                p = paragraph._element
                p.getparent().remove(p)

    @staticmethod
    def dejar_el_marcador(document):
        
        valores = {
            '${eliminar}': ''
        }
        
        # Iterar sobre todos los párrafos del documento y reemplazar los marcadores
        for paragraph in document.paragraphs:
            for marcador, reemplazo in valores.items():
                if marcador in paragraph.text:  # Verifica si el marcador está en el párrafo
                    # Reemplaza el marcador con el texto proporcionado
                    paragraph.text = paragraph.text.replace(marcador, reemplazo)
        
    # Función para eliminar los párrafos que contienen ciertos marcadores
    @staticmethod
    def eliminar_parrafos_innecesarios(document):
        for paragraph in document.paragraphs:
            if '${texto_1}' in paragraph.text or '${texto_2}' in paragraph.text or '${texto_3}' in paragraph.text or '${texto_6}' in paragraph.text:
                p = paragraph._element
                p.getparent().remove(p)

    @staticmethod
    def eliminar_parrafos_innecesarios_fractioned(document):
        for paragraph in document.paragraphs:
            if '${texto_1}' in paragraph.text:
                p = paragraph._element               
                p.getparent().remove(p)
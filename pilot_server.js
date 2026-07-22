const http = require('http');
const fs = require('fs');
const path = require('path');
const os = require('os');
const { spawn } = require('child_process');

const root = __dirname;
const port = Number(process.env.PORT || 8787);
const htriExcelCopyPs1Base64 = '77u/cGFyYW0oCiAgW1BhcmFtZXRlcihNYW5kYXRvcnk9JHRydWUpXVtzdHJpbmddJE1hbmlmZXN0UGF0aCwKICBbUGFyYW1ldGVyKE1hbmRhdG9yeT0kdHJ1ZSldW3N0cmluZ10kT3V0cHV0RGlyLAogIFtQYXJhbWV0ZXIoTWFuZGF0b3J5PSR0cnVlKV1bc3RyaW5nXSRUZW1wbGF0ZURpcgopCgokRXJyb3JBY3Rpb25QcmVmZXJlbmNlID0gJ1N0b3AnCiRtYW5pZmVzdCA9IEdldC1Db250ZW50IC1MaXRlcmFsUGF0aCAkTWFuaWZlc3RQYXRoIC1SYXcgLUVuY29kaW5nIFVURjggfCBDb252ZXJ0RnJvbS1Kc29uCk5ldy1JdGVtIC1JdGVtVHlwZSBEaXJlY3RvcnkgLUZvcmNlIC1QYXRoICRPdXRwdXREaXIgfCBPdXQtTnVsbAoKZnVuY3Rpb24gU2FmZS1QYXJ0KFtzdHJpbmddJHRleHQsIFtzdHJpbmddJGZhbGxiYWNrKSB7CiAgJHZhbHVlID0gaWYgKFtzdHJpbmddOjpJc051bGxPcldoaXRlU3BhY2UoJHRleHQpKSB7ICRmYWxsYmFjayB9IGVsc2UgeyAkdGV4dCB9CiAgcmV0dXJuICgkdmFsdWUgLXJlcGxhY2UgJ1tcLzoqPyI8PnxdJywgJ18nIC1yZXBsYWNlICdccysnLCAnJykKfQoKZnVuY3Rpb24gVW5pcXVlLVBhdGgoW3N0cmluZ10kcGF0aCkgewogIGlmICghKFRlc3QtUGF0aCAtTGl0ZXJhbFBhdGggJHBhdGgpKSB7IHJldHVybiAkcGF0aCB9CiAgJGRpciA9IFtJTy5QYXRoXTo6R2V0RGlyZWN0b3J5TmFtZSgkcGF0aCkKICAkYmFzZSA9IFtJTy5QYXRoXTo6R2V0RmlsZU5hbWVXaXRob3V0RXh0ZW5zaW9uKCRwYXRoKQogICRleHQgPSBbSU8uUGF0aF06OkdldEV4dGVuc2lvbigkcGF0aCkKICAkbiA9IDIKICBkbyB7CiAgICAkY2FuZGlkYXRlID0gSm9pbi1QYXRoICRkaXIgKCRiYXNlICsgJ18nICsgJG4gKyAkZXh0KQogICAgJG4rKwogIH0gd2hpbGUgKFRlc3QtUGF0aCAtTGl0ZXJhbFBhdGggJGNhbmRpZGF0ZSkKICByZXR1cm4gJGNhbmRpZGF0ZQp9CgpmdW5jdGlvbiBTaGVldC1MaWtlKCR3b3JrYm9vaywgW3N0cmluZ10kcGF0dGVybikgewogIGZvcmVhY2ggKCR3cyBpbiBAKCR3b3JrYm9vay5Xb3Jrc2hlZXRzKSkgewogICAgaWYgKCR3cy5OYW1lIC1tYXRjaCAkcGF0dGVybikgeyByZXR1cm4gJHdzIH0KICB9CiAgcmV0dXJuICRudWxsCn0KCmZ1bmN0aW9uIFQoJHNoZWV0LCBbc3RyaW5nXSRhZGRyKSB7CiAgdHJ5IHsgcmV0dXJuIFtzdHJpbmddJHNoZWV0LlJhbmdlKCRhZGRyKS5UZXh0IH0gY2F0Y2ggeyByZXR1cm4gJycgfQp9CgpmdW5jdGlvbiBQdXQoJHNoZWV0LCBbc3RyaW5nXSRhZGRyLCAkdmFsdWUpIHsKICBpZiAoJG51bGwgLWVxICR2YWx1ZSkgeyAkdmFsdWUgPSAnJyB9CiAgJHNoZWV0LlJhbmdlKCRhZGRyKS5WYWx1ZTIgPSAkdmFsdWUKfQoKZnVuY3Rpb24gUHV0RGFzaCgkc2hlZXQsIFtzdHJpbmddJGFkZHIsICR2YWx1ZSkgewogICRzID0gKFtzdHJpbmddJHZhbHVlKS5UcmltKCkKICBpZiAoW3N0cmluZ106OklzTnVsbE9yV2hpdGVTcGFjZSgkcykgLW9yICRzIC1lcSAnLycpIHsgJHMgPSAnLScgfQogICRzaGVldC5SYW5nZSgkYWRkcikuVmFsdWUyID0gJHMKfQoKZnVuY3Rpb24gUHV0RGFzaE51bWJlcigkc2hlZXQsIFtzdHJpbmddJGFkZHIsICR2YWx1ZSkgewogIFB1dERhc2ggJHNoZWV0ICRhZGRyICR2YWx1ZQogIGlmICgoW3N0cmluZ10kc2hlZXQuUmFuZ2UoJGFkZHIpLlZhbHVlMikgLW5lICctJykgewogICAgdHJ5IHsgJHNoZWV0LlJhbmdlKCRhZGRyKS5OdW1iZXJGb3JtYXQgPSAnMC4jIyMnIH0gY2F0Y2gge30KICB9Cn0KCmZ1bmN0aW9uIEFkZHIoW2Jvb2xdJGNvbmRpdGlvbiwgW3N0cmluZ10kd2hlblRydWUsIFtzdHJpbmddJHdoZW5GYWxzZSkgewogIGlmICgkY29uZGl0aW9uKSB7IHJldHVybiAkd2hlblRydWUgfQogIHJldHVybiAkd2hlbkZhbHNlCn0KCmZ1bmN0aW9uIEZpcnN0LU5vbkJsYW5rKCR2YWx1ZXMpIHsKICBmb3JlYWNoICgkdiBpbiAkdmFsdWVzKSB7CiAgICBpZiAoJG51bGwgLW5lICR2IC1hbmQgLW5vdCBbc3RyaW5nXTo6SXNOdWxsT3JXaGl0ZVNwYWNlKFtzdHJpbmddJHYpKSB7IHJldHVybiAkdiB9CiAgfQogIHJldHVybiAnJwp9CgpmdW5jdGlvbiBGaXJzdC1Vc2VmdWwoJHZhbHVlcykgewogIGZvcmVhY2ggKCR2IGluICR2YWx1ZXMpIHsKICAgICRzID0gW3N0cmluZ10kdgogICAgaWYgKC1ub3QgW3N0cmluZ106OklzTnVsbE9yV2hpdGVTcGFjZSgkcykgLWFuZCAkcyAtbm90bWF0Y2ggJ14oXC98U2VydmljZXxDbGVhbnxcKikkJykgeyByZXR1cm4gJHYgfQogIH0KICByZXR1cm4gJycKfQoKZnVuY3Rpb24gTnVtKCR2YWx1ZSkgewogICRzID0gKFtzdHJpbmddJHZhbHVlKS5UcmltKCkgLXJlcGxhY2UgJywnLCAnJwogICRuID0gMC4wCiAgaWYgKFtkb3VibGVdOjpUcnlQYXJzZSgkcywgW0dsb2JhbGl6YXRpb24uTnVtYmVyU3R5bGVzXTo6RmxvYXQsIFtHbG9iYWxpemF0aW9uLkN1bHR1cmVJbmZvXTo6SW52YXJpYW50Q3VsdHVyZSwgW3JlZl0kbikpIHsgcmV0dXJuICRuIH0KICByZXR1cm4gJG51bGwKfQoKZnVuY3Rpb24gRmlyc3QtTnVtYmVyVGV4dCgkdmFsdWVzKSB7CiAgZm9yZWFjaCAoJHYgaW4gJHZhbHVlcykgewogICAgJHMgPSAoW3N0cmluZ10kdikuVHJpbSgpCiAgICBpZiAoW3N0cmluZ106OklzTnVsbE9yV2hpdGVTcGFjZSgkcykpIHsgY29udGludWUgfQogICAgaWYgKCRudWxsIC1uZSAoTnVtICRzKSkgeyByZXR1cm4gJHMgfQogIH0KICByZXR1cm4gJycKfQoKZnVuY3Rpb24gQWRkLU9uZSgkdmFsdWUpIHsKICAkbiA9IE51bSAkdmFsdWUKICBpZiAoJG51bGwgLWVxICRuKSB7IHJldHVybiAkdmFsdWUgfQogIHJldHVybiBbaW50XSgkbiArIDEpCn0KCmZ1bmN0aW9uIE1heC1OdW1iZXIoJHZhbHVlcywgW2RvdWJsZV0kZGVmYXVsdCkgewogICRtYXggPSAkZGVmYXVsdAogIGZvcmVhY2ggKCR2IGluICR2YWx1ZXMpIHsKICAgICRuID0gTnVtICR2CiAgICBpZiAoJG51bGwgLW5lICRuIC1hbmQgJG4gLWd0ICRtYXgpIHsgJG1heCA9ICRuIH0KICB9CiAgcmV0dXJuICRtYXgKfQoKZnVuY3Rpb24gTWluLU51bWJlcigkdmFsdWVzLCBbZG91YmxlXSRkZWZhdWx0KSB7CiAgJG1pbiA9ICRkZWZhdWx0CiAgZm9yZWFjaCAoJHYgaW4gJHZhbHVlcykgewogICAgJG4gPSBOdW0gJHYKICAgIGlmICgkbnVsbCAtbmUgJG4gLWFuZCAkbiAtbHQgJG1pbikgeyAkbWluID0gJG4gfQogIH0KICByZXR1cm4gJG1pbgp9CgpmdW5jdGlvbiBNb2RlbC1Gcm9tTmFtZShbc3RyaW5nXSRuYW1lKSB7CiAgJG0gPSBbcmVnZXhdOjpNYXRjaCgkbmFtZSwgJyhaUzAyNXxLQzEyfE1cZHsyLDN9W0EtWl0qfE1YXGQrW0EtWl0qfE1DXGQrfE1TXGQrfE1BXGQrKScsICdJZ25vcmVDYXNlJykKICBpZiAoJG0uU3VjY2VzcykgeyByZXR1cm4gJG0uVmFsdWUuVG9VcHBlcigpIH0KICByZXR1cm4gJycKfQoKZnVuY3Rpb24gUGxhdGVzLUZyb21OYW1lKFtzdHJpbmddJG5hbWUpIHsKICAkbSA9IFtyZWdleF06Ok1hdGNoKCRuYW1lLCAnKFxkKylccypwbCcsICdJZ25vcmVDYXNlJykKICBpZiAoJG0uU3VjY2VzcykgeyByZXR1cm4gJG0uR3JvdXBzWzFdLlZhbHVlIH0KICByZXR1cm4gJycKfQoKZnVuY3Rpb24gQ2xlYXItS25vd25TbGFzaENlbGxzKCRzaGVldCwgW2Jvb2xdJGlzR3BoZSkgewogICRjZWxscyA9IEAoJ0QxMicsJ0gxMicpCiAgZm9yZWFjaCAoJGFkZHIgaW4gJGNlbGxzKSB7IFB1dCAkc2hlZXQgJGFkZHIgJycgfQp9CgpmdW5jdGlvbiBPdmVycmlkZS1WYWx1ZSgkaXRlbSwgW3N0cmluZ10ka2V5LCAkZmFsbGJhY2sgPSAnJykgewogIGlmICgkbnVsbCAtbmUgJGl0ZW0ub3ZlcnJpZGVzKSB7CiAgICAkcHJvcCA9ICRpdGVtLm92ZXJyaWRlcy5QU09iamVjdC5Qcm9wZXJ0aWVzWyRrZXldCiAgICBpZiAoJG51bGwgLW5lICRwcm9wIC1hbmQgLW5vdCBbc3RyaW5nXTo6SXNOdWxsT3JXaGl0ZVNwYWNlKFtzdHJpbmddJHByb3AuVmFsdWUpKSB7IHJldHVybiAkcHJvcC5WYWx1ZSB9CiAgfQogIHJldHVybiAkZmFsbGJhY2sKfQoKZnVuY3Rpb24gVW5pdC1UZXh0KFtzdHJpbmddJHVuaXQpIHsKICAkdSA9IChbc3RyaW5nXSR1bml0KS5UcmltKCkKICBpZiAoW3N0cmluZ106OklzTnVsbE9yV2hpdGVTcGFjZSgkdSkpIHsgcmV0dXJuICcnIH0KICBpZiAoJHUuU3RhcnRzV2l0aCgnKCcpIC1hbmQgJHUuRW5kc1dpdGgoJyknKSkgeyByZXR1cm4gJHUgfQogIHJldHVybiAnKCcgKyAkdSArICcpJwp9CgpmdW5jdGlvbiBTcGxpdC1QYWlyKFtzdHJpbmddJHRleHQpIHsKICAkcGFydHMgPSAoW3N0cmluZ10kdGV4dCkgLXNwbGl0ICdccyovXHMqfFxzKixccyonCiAgaWYgKCRwYXJ0cy5Db3VudCAtZ2UgMikgeyByZXR1cm4gQCgkcGFydHNbMF0uVHJpbSgpLCAkcGFydHNbMV0uVHJpbSgpKSB9CiAgcmV0dXJuIEAoKFtzdHJpbmddJHRleHQpLlRyaW0oKSwgJycpCn0KCmZ1bmN0aW9uIFNldC1Nb2RlbFByZWZpeChbc3RyaW5nXSRtb2RlbCwgW3N0cmluZ10kbWF0ZXJpYWwpIHsKICAkbSA9IChbc3RyaW5nXSRtb2RlbCkuVHJpbSgpCiAgaWYgKCRtIC1ub3RtYXRjaCAnXihNQ3xNUyknKSB7IHJldHVybiAkbSB9CiAgaWYgKCRtYXRlcmlhbCAtbWF0Y2ggJ1N0YWlubGVzc3xTVVN8Xk1TJCcpIHsgcmV0dXJuICgkbSAtcmVwbGFjZSAnXihNQ3xNUyknLCAnTVMnKSB9CiAgaWYgKCRtYXRlcmlhbCAtbWF0Y2ggJ0NvcHBlcnxeTUMkJykgeyByZXR1cm4gKCRtIC1yZXBsYWNlICdeKE1DfE1TKScsICdNQycpIH0KICByZXR1cm4gJG0KfQoKZnVuY3Rpb24gRGlzcGxheS1Tb2xkZXIoW3N0cmluZ10kbWF0ZXJpYWwsIFtzdHJpbmddJG1vZGVsKSB7CiAgaWYgKCRtYXRlcmlhbCAtbWF0Y2ggJ1N0YWlubGVzc3xTVVN8Xk1TJCcpIHsgcmV0dXJuICdTdGFpbmxlc3MgU3RlZWwnIH0KICBpZiAoJG1hdGVyaWFsIC1tYXRjaCAnQ29wcGVyfF5NQyQnKSB7IHJldHVybiAnQ29wcGVyJyB9CiAgaWYgKCRtb2RlbCAtbWF0Y2ggJ15NUycpIHsgcmV0dXJuICdTdGFpbmxlc3MgU3RlZWwnIH0KICByZXR1cm4gJ0NvcHBlcicKfQoKZnVuY3Rpb24gR3BoZS1Db25uZWN0aW9uVGV4dChbc3RyaW5nXSRtb2RlbCkgewogICRzaXplID0gJzUwQScKICBpZiAoJG1vZGVsIC1tYXRjaCAnXk03ME0kJykgeyAkc2l6ZSA9ICc2NUEnIH0KICBlbHNlaWYgKCRtb2RlbCAtbWF0Y2ggJ15NKFxkezIsM30pW0EtWl0qJCcpIHsgJHNpemUgPSAkbWF0Y2hlc1sxXSArICdBJyB9CiAgcmV0dXJuICJLUyAxMGsgJHNpemUgTWV0YWwgTGluZWQgV2l0aCBTVFMzMTYgKFN0dWQgVHlwZSkiCn0KCmZ1bmN0aW9uIEhlYXQtRmFjdG9yKFtzdHJpbmddJHVuaXQpIHsKICAkdSA9ICgoW3N0cmluZ10kdW5pdCkuVHJpbSgnKCcsJyknLCcgJykgLXJlcGxhY2UgJ1xzKycsICcnKS5Ub0xvd2VySW52YXJpYW50KCkKICBzd2l0Y2ggLVJlZ2V4ICgkdSkgewogICAgJ15rdyQnIHsgcmV0dXJuIDEwMDAuMCB9CiAgICAnXnckJyB7IHJldHVybiAxLjAgfQogICAgJ14oa2NhbC9ocnxrY2FsaHxrY2FsL2hvdXIpJCcgeyByZXR1cm4gNDE4Ni44IC8gMzYwMC4wIH0KICAgICdeKGJ0dS9ocnxidHVofGJ0dS9ob3VyKSQnIHsgcmV0dXJuIDAuMjkzMDcxMDcgfQogICAgZGVmYXVsdCB7IHJldHVybiAkbnVsbCB9CiAgfQp9CgpmdW5jdGlvbiBQcmVzc3VyZS1GYWN0b3IoW3N0cmluZ10kdW5pdCkgewogICR1ID0gKChbc3RyaW5nXSR1bml0KS5UcmltKCcoJywnKScsJyAnKSAtcmVwbGFjZSAnXHMrJywgJycgLXJlcGxhY2UgJ18nLCAnJyAtcmVwbGFjZSAnLScsICcnKS5Ub0xvd2VySW52YXJpYW50KCkKICBpZiAoJHUgLW1hdGNoICdeKHBhfGtwYXxtcGF8YmFyfGtnZi9jbTJ8a2dmY20yfGF0bXxwc2l8dG9ycikoZ3xhKSQnKSB7ICR1ID0gJG1hdGNoZXNbMV0gfQogIHN3aXRjaCAtUmVnZXggKCR1KSB7CiAgICAnXnBhJCcgeyByZXR1cm4gMS4wIH0KICAgICdea3BhJCcgeyByZXR1cm4gMTAwMC4wIH0KICAgICdebXBhJCcgeyByZXR1cm4gMTAwMDAwMC4wIH0KICAgICdeYmFyJCcgeyByZXR1cm4gMTAwMDAwLjAgfQogICAgJ15rZ2YvY20yJHxea2dmY20yJCcgeyByZXR1cm4gOTgwNjYuNSB9CiAgICAnXmF0bSQnIHsgcmV0dXJuIDEwMTMyNS4wIH0KICAgICdecHNpJCcgeyByZXR1cm4gNjg5NC43NTcyOTMxNjggfQogICAgJ150b3JyJCcgeyByZXR1cm4gMTMzLjMyMjM2ODQyMSB9CiAgICBkZWZhdWx0IHsgcmV0dXJuICRudWxsIH0KICB9Cn0KCmZ1bmN0aW9uIENvbnZlcnQtVW5pdFZhbHVlKCR2YWx1ZSwgW3N0cmluZ10kZnJvbVVuaXQsIFtzdHJpbmddJHRvVW5pdCwgW3N0cmluZ10ka2luZCkgewogICRuID0gTnVtICR2YWx1ZQogIGlmICgkbnVsbCAtZXEgJG4pIHsgcmV0dXJuICR2YWx1ZSB9CiAgJGZyb20gPSBpZiAoJGtpbmQgLWVxICdoZWF0JykgeyBIZWF0LUZhY3RvciAkZnJvbVVuaXQgfSBlbHNlIHsgUHJlc3N1cmUtRmFjdG9yICRmcm9tVW5pdCB9CiAgJHRvID0gaWYgKCRraW5kIC1lcSAnaGVhdCcpIHsgSGVhdC1GYWN0b3IgJHRvVW5pdCB9IGVsc2UgeyBQcmVzc3VyZS1GYWN0b3IgJHRvVW5pdCB9CiAgaWYgKCRudWxsIC1lcSAkZnJvbSAtb3IgJG51bGwgLWVxICR0byAtb3IgJHRvIC1lcSAwKSB7IHJldHVybiAkdmFsdWUgfQogIHJldHVybiBbTWF0aF06OlJvdW5kKCgkbiAqICRmcm9tIC8gJHRvKSwgNikKfQoKZnVuY3Rpb24gQ29udmVydC1DZWxscygkc2hlZXQsIFtzdHJpbmdbXV0kY2VsbHMsIFtzdHJpbmddJHVuaXRDZWxsLCBbc3RyaW5nXSR0YXJnZXRVbml0LCBbc3RyaW5nXSRraW5kKSB7CiAgaWYgKFtzdHJpbmddOjpJc051bGxPcldoaXRlU3BhY2UoJHRhcmdldFVuaXQpKSB7IHJldHVybiB9CiAgJGZyb21Vbml0ID0gVCAkc2hlZXQgJHVuaXRDZWxsCiAgZm9yZWFjaCAoJGFkZHIgaW4gJGNlbGxzKSB7CiAgICBQdXQgJHNoZWV0ICRhZGRyIChDb252ZXJ0LVVuaXRWYWx1ZSAoVCAkc2hlZXQgJGFkZHIpICRmcm9tVW5pdCAkdGFyZ2V0VW5pdCAka2luZCkKICAgIHRyeSB7ICRzaGVldC5SYW5nZSgkYWRkcikuTnVtYmVyRm9ybWF0ID0gJzAuIyMjIyMjJyB9IGNhdGNoIHt9CiAgfQogIFB1dCAkc2hlZXQgJHVuaXRDZWxsIChVbml0LVRleHQgJHRhcmdldFVuaXQpCn0KCmZ1bmN0aW9uIEFwcGx5LUh0cmlPdmVycmlkZXMoJGRzLCAkaXRlbSwgW2Jvb2xdJGlzR3BoZSkgewogIFB1dCAkZHMgJ0IxJyAoT3ZlcnJpZGUtVmFsdWUgJGl0ZW0gJ0N1c3RvbWVyJyAoVCAkZHMgJ0IxJykpCiAgUHV0ICRkcyAnQjInIChPdmVycmlkZS1WYWx1ZSAkaXRlbSAnUHJvamVjdCBOYW1lJyAoVCAkZHMgJ0IyJykpCiAgUHV0ICRkcyAnQjMnIChPdmVycmlkZS1WYWx1ZSAkaXRlbSAnQ29udGFjdCcgKFQgJGRzICdCMycpKQogIFB1dCAkZHMgJ0ozJyAoT3ZlcnJpZGUtVmFsdWUgJGl0ZW0gJ0l0ZW0gTm8uJyAoVCAkZHMgJ0ozJykpCiAgUHV0ICRkcyAnQjQnIChPdmVycmlkZS1WYWx1ZSAkaXRlbSAnU2VydmljZScgKFQgJGRzICdCNCcpKQogIFB1dCAkZHMgJ0MxMScgKE92ZXJyaWRlLVZhbHVlICRpdGVtICdNZWRpYSBIb3QnIChUICRkcyAnQzExJykpCiAgUHV0ICRkcyAnRzExJyAoT3ZlcnJpZGUtVmFsdWUgJGl0ZW0gJ01lZGlhIENvbGQnIChUICRkcyAnRzExJykpCgogICRoZWF0VmFsdWUgPSBPdmVycmlkZS1WYWx1ZSAkaXRlbSAnSGVhdCBjYXBhY2l0eScgJycKICBpZiAoLW5vdCBbc3RyaW5nXTo6SXNOdWxsT3JXaGl0ZVNwYWNlKFtzdHJpbmddJGhlYXRWYWx1ZSkpIHsgUHV0ICRkcyAnRTEyJyAkaGVhdFZhbHVlIH0KICAkaGVhdFVuaXQgPSBPdmVycmlkZS1WYWx1ZSAkaXRlbSAnSGVhdCBjYXBhY2l0eSBVbml0JyAnJwogIGlmICgtbm90IFtzdHJpbmddOjpJc051bGxPcldoaXRlU3BhY2UoW3N0cmluZ10kaGVhdFVuaXQpKSB7IFB1dCAkZHMgJ0oxMicgKFVuaXQtVGV4dCAkaGVhdFVuaXQpIH0KICBDb252ZXJ0LUNlbGxzICRkcyBAKCdFMTInKSAnSjEyJyAoT3ZlcnJpZGUtVmFsdWUgJGl0ZW0gJ0hlYXQgY2FwYWNpdHkgVGFyZ2V0IFVuaXQnICcnKSAnaGVhdCcKCiAgJHBkSG90ID0gQWRkciAkaXNHcGhlICdDMTYnICdDMTcnCiAgJHBkQ29sZCA9IEFkZHIgJGlzR3BoZSAnRzE2JyAnRzE3JwogICRwZFVuaXQgPSBBZGRyICRpc0dwaGUgJ0oxNicgJ0oxNycKICBQdXQgJGRzICRwZEhvdCAoT3ZlcnJpZGUtVmFsdWUgJGl0ZW0gJ1ByZXNzdXJlIGRyb3AgSG90JyAoVCAkZHMgJHBkSG90KSkKICBQdXQgJGRzICRwZENvbGQgKE92ZXJyaWRlLVZhbHVlICRpdGVtICdQcmVzc3VyZSBkcm9wIENvbGQnIChUICRkcyAkcGRDb2xkKSkKICAkcGRVbml0VmFsdWUgPSBPdmVycmlkZS1WYWx1ZSAkaXRlbSAnUHJlc3N1cmUgZHJvcCBVbml0JyAnJwogIGlmICgtbm90IFtzdHJpbmddOjpJc051bGxPcldoaXRlU3BhY2UoW3N0cmluZ10kcGRVbml0VmFsdWUpKSB7IFB1dCAkZHMgJHBkVW5pdCAoVW5pdC1UZXh0ICRwZFVuaXRWYWx1ZSkgfQogIENvbnZlcnQtQ2VsbHMgJGRzIEAoJHBkSG90LCAkcGRDb2xkKSAkcGRVbml0IChPdmVycmlkZS1WYWx1ZSAkaXRlbSAnUHJlc3N1cmUgZHJvcCBUYXJnZXQgVW5pdCcgJycpICdwcmVzc3VyZScKCiAgJG9wSG90ID0gQWRkciAkaXNHcGhlICdDMTknICdDMTgnCiAgJG9wQ29sZCA9IEFkZHIgJGlzR3BoZSAnRzE5JyAnRzE4JwogICRvcFVuaXQgPSBBZGRyICRpc0dwaGUgJ0oxOScgJ0oxOCcKICBQdXQgJGRzICRvcEhvdCAoT3ZlcnJpZGUtVmFsdWUgJGl0ZW0gJ09wZXJhdGluZyBQcmVzc3VyZSBIb3QnIChUICRkcyAkb3BIb3QpKQogIFB1dCAkZHMgJG9wQ29sZCAoT3ZlcnJpZGUtVmFsdWUgJGl0ZW0gJ09wZXJhdGluZyBQcmVzc3VyZSBDb2xkJyAoVCAkZHMgJG9wQ29sZCkpCiAgJG9wVW5pdFZhbHVlID0gT3ZlcnJpZGUtVmFsdWUgJGl0ZW0gJ09wZXJhdGluZyBQcmVzc3VyZSBVbml0JyAnJwogIGlmICgtbm90IFtzdHJpbmddOjpJc051bGxPcldoaXRlU3BhY2UoW3N0cmluZ10kb3BVbml0VmFsdWUpKSB7IFB1dCAkZHMgJG9wVW5pdCAoVW5pdC1UZXh0ICRvcFVuaXRWYWx1ZSkgfQogIENvbnZlcnQtQ2VsbHMgJGRzIEAoJG9wSG90LCAkb3BDb2xkKSAkb3BVbml0IChPdmVycmlkZS1WYWx1ZSAkaXRlbSAnT3BlcmF0aW5nIFByZXNzdXJlIFRhcmdldCBVbml0JyAnJykgJ3ByZXNzdXJlJwoKICAkZHQgPSBTcGxpdC1QYWlyIChPdmVycmlkZS1WYWx1ZSAkaXRlbSAnRGVzaWduIFRlbXBlcmF0dXJlJyAnJykKICBpZiAoLW5vdCBbc3RyaW5nXTo6SXNOdWxsT3JXaGl0ZVNwYWNlKCRkdFswXSkpIHsKICAgIFB1dCAkZHMgKEFkZHIgJGlzR3BoZSAnRDQ1JyAnQzQwJykgJGR0WzBdCiAgICBpZiAoLW5vdCBbc3RyaW5nXTo6SXNOdWxsT3JXaGl0ZVNwYWNlKCRkdFsxXSkpIHsgUHV0ICRkcyAoQWRkciAkaXNHcGhlICdINDUnICdHNDAnKSAkZHRbMV0gfQogIH0KICAkZHRVbml0ID0gT3ZlcnJpZGUtVmFsdWUgJGl0ZW0gJ0Rlc2lnbiBUZW1wZXJhdHVyZSBVbml0JyAnJwogIGlmICgtbm90IFtzdHJpbmddOjpJc051bGxPcldoaXRlU3BhY2UoW3N0cmluZ10kZHRVbml0KSkgeyBQdXQgJGRzIChBZGRyICRpc0dwaGUgJ0o0NScgJ0o0MCcpIChVbml0LVRleHQgJGR0VW5pdCkgfQoKICAkZHAgPSBTcGxpdC1QYWlyIChPdmVycmlkZS1WYWx1ZSAkaXRlbSAnRGVzaWduIC8gVGVzdCBQcmVzc3VyZScgJycpCiAgaWYgKC1ub3QgW3N0cmluZ106OklzTnVsbE9yV2hpdGVTcGFjZSgkZHBbMF0pKSB7CiAgICBQdXQgJGRzIChBZGRyICRpc0dwaGUgJ0U0NicgJ0M0MScpICRkcFswXQogICAgaWYgKC1ub3QgW3N0cmluZ106OklzTnVsbE9yV2hpdGVTcGFjZSgkZHBbMV0pKSB7IFB1dCAkZHMgKEFkZHIgJGlzR3BoZSAnRzQ2JyAnRzQxJykgJGRwWzFdIH0KICB9CiAgJGRlc2lnblVuaXQgPSBPdmVycmlkZS1WYWx1ZSAkaXRlbSAnRGVzaWduIC8gVGVzdCBQcmVzc3VyZSBVbml0JyAnJwogIGlmICgtbm90IFtzdHJpbmddOjpJc051bGxPcldoaXRlU3BhY2UoW3N0cmluZ10kZGVzaWduVW5pdCkpIHsgUHV0ICRkcyAoQWRkciAkaXNHcGhlICdKNDYnICdKNDEnKSAoVW5pdC1UZXh0ICRkZXNpZ25Vbml0KSB9CiAgQ29udmVydC1DZWxscyAkZHMgQCgoQWRkciAkaXNHcGhlICdFNDYnICdDNDEnKSwgKEFkZHIgJGlzR3BoZSAnRzQ2JyAnRzQxJykpIChBZGRyICRpc0dwaGUgJ0o0NicgJ0o0MScpIChPdmVycmlkZS1WYWx1ZSAkaXRlbSAnRGVzaWduIC8gVGVzdCBQcmVzc3VyZSBUYXJnZXQgVW5pdCcgJycpICdwcmVzc3VyZScKCiAgaWYgKCRpc0dwaGUpIHsKICAgIFB1dCAkZHMgJ0M0MCcgKE92ZXJyaWRlLVZhbHVlICRpdGVtICdHYXNrZXQgTWF0ZXJpYWwnIChUICRkcyAnQzQwJykpCiAgfQp9CgpmdW5jdGlvbiBBcHBseS1IdHJpVmFsdWVzKCR0YXJnZXQsICRzb3VyY2UsICRpdGVtLCBbc3RyaW5nXSRraW5kKSB7CiAgJGRzID0gU2hlZXQtTGlrZSAkdGFyZ2V0ICdNQyBEYXRhc2hlZXR8RGF0YXNoZWV0JwogICRhcGkgPSBTaGVldC1MaWtlICRzb3VyY2UgJ0FQSVxzKjY2MicKICAkb3V0ID0gU2hlZXQtTGlrZSAkc291cmNlICdPdXRwdXRccypTdW1tYXJ5JwogICRmaW4gPSBTaGVldC1MaWtlICRzb3VyY2UgJ0ZpbmFsXHMqUmVzdWx0cycKICBpZiAoJG51bGwgLWVxICRkcykgeyB0aHJvdyAnVGVtcGxhdGUgTUMgRGF0YXNoZWV0IHNoZWV0IG5vdCBmb3VuZC4nIH0KICBpZiAoJG51bGwgLWVxICRhcGkpIHsgdGhyb3cgJ1NvdXJjZSBBUEkgNjYyIHNoZWV0IG5vdCBmb3VuZC4nIH0KCiAgdHJ5IHsgJGRzLlVzZWRSYW5nZS5WYWx1ZTIgPSAkZHMuVXNlZFJhbmdlLlZhbHVlMiB9IGNhdGNoIHt9CgogICRiYXNlID0gW0lPLlBhdGhdOjpHZXRGaWxlTmFtZVdpdGhvdXRFeHRlbnNpb24oJGl0ZW0ubmFtZSkKICAkbW9kZWwgPSBGaXJzdC1Ob25CbGFuayBAKChNb2RlbC1Gcm9tTmFtZSAkYmFzZSksIChUICRhcGkgJ0wxMCcpLCAoVCAkZmluICdMMTEnKSkKICAkbWF0ZXJpYWxPdmVycmlkZSA9IE92ZXJyaWRlLVZhbHVlICRpdGVtICdTb2xkZXJpbmcgTWF0ZXJpYWwnICcnCiAgJG1vZGVsID0gU2V0LU1vZGVsUHJlZml4ICRtb2RlbCAkbWF0ZXJpYWxPdmVycmlkZQogICRwbGF0ZXMgPSBGaXJzdC1Ob25CbGFuayBAKChQbGF0ZXMtRnJvbU5hbWUgJGJhc2UpLCAoVCAkYXBpICdMMzgnKSkKICAkaXNHcGhlID0gJGtpbmQgLWVxICdHUEhFJwoKICBQdXQgJGRzICdCMScgKFQgJGFwaSAnRDYnKQogIFB1dCAkZHMgJ0IyJyAnJwogIFB1dCAkZHMgJ0IzJyAnJwogIFB1dCAkZHMgJ0ozJyAoVCAkYXBpICdNOScpCiAgUHV0ICRkcyAnQjQnIChUICRhcGkgJ0Q5JykKICBQdXQgJGRzICdKNCcgKEdldC1EYXRlIC1Gb3JtYXQgJ3l5eXktTU0tZGQnKQogIFB1dCAkZHMgJ0M2JyAoIjEgeCAkbW9kZWwgLSAkcGxhdGVzIENvdW50ZXJjdXJyZW50IEZsb3ciKQogIFB1dCAkZHMgJ0E4JyAnICBUaGVybWFsIGRhdGEgZm9yIDEgdW5pdChzKSBpbiBwYXJhbGxlbCBhbmQgMSB1bml0KHMpIGluIHNlcmllcycKCiAgUHV0ICRkcyAnQzExJyAoVCAkYXBpICdMMTQnKQogIFB1dCAkZHMgJ0cxMScgKFQgJGFwaSAnVjE0JykKICBQdXQgJGRzICdFMTInIChGaXJzdC1OdW1iZXJUZXh0IEAoKFQgJGZpbiAnSTMwJyksIChUICRvdXQgJ1MyNicpLCAoVCAkb3V0ICdUMjYnKSwgKFQgJGFwaSAnTDQwJykpKQogIFB1dCAkZHMgJ0QxMicgJycKICBQdXQgJGRzICdIMTInICcnCiAgUHV0ICRkcyAnSjEyJyAnKGtjYWwvaHIpJwogIFB1dCAkZHMgJ0MxMycgKFQgJGFwaSAnTDE2JykKICBQdXQgJGRzICdHMTMnIChUICRhcGkgJ1YxNicpCiAgUHV0ICRkcyAnSjEzJyAoJygnICsgKFQgJGFwaSAnSTE2JykgKyAnKScpCiAgUHV0ICRkcyAnQzE0JyAoRmlyc3QtTnVtYmVyVGV4dCBAKChUICRvdXQgJ0cxNCcpLCAoVCAkb3V0ICdIMTQnKSwgKFQgJGZpbiAnSTE0JyksIChUICRmaW4gJ0oxNCcpLCAoVCAkYXBpICdNMjInKSkpCiAgUHV0ICRkcyAnQzE1JyAoRmlyc3QtTnVtYmVyVGV4dCBAKChUICRvdXQgJ00xNCcpLCAoVCAkb3V0ICdOMTQnKSwgKFQgJGZpbiAnTTE0JyksIChUICRmaW4gJ04xNCcpLCAoVCAkYXBpICdSMjInKSkpCiAgUHV0ICRkcyAnRzE0JyAoRmlyc3QtTnVtYmVyVGV4dCBAKChUICRvdXQgJ1ExNCcpLCAoVCAkb3V0ICdSMTQnKSwgKFQgJGZpbiAnTzE0JyksIChUICRmaW4gJ1AxNCcpLCAoVCAkYXBpICdXMjInKSkpCiAgUHV0ICRkcyAnRzE1JyAoRmlyc3QtTnVtYmVyVGV4dCBAKChUICRvdXQgJ1MxNCcpLCAoVCAkb3V0ICdUMTQnKSwgKFQgJGZpbiAnUTE0JyksIChUICRmaW4gJ1IxNCcpLCAoVCAkYXBpICdBQjIyJykpKQogIFB1dCAkZHMgJ0oxNCcgKCcoJyArIChUICRhcGkgJ0kyMicpICsgJyknKQogIFB1dCAkZHMgJ0oxNScgKCcoJyArIChUICRhcGkgJ0kyMicpICsgJyknKQoKICAkdmFwSG90SW4gPSBGaXJzdC1OdW1iZXJUZXh0IEAoKFQgJG91dCAnRzE1JyksIChUICRvdXQgJ0gxNScpLCAoVCAkZmluICdJMTMnKSwgKFQgJGZpbiAnSjEzJykpCiAgJHZhcEhvdE91dCA9IEZpcnN0LU51bWJlclRleHQgQCgoVCAkb3V0ICdNMTUnKSwgKFQgJG91dCAnTjE1JyksIChUICRmaW4gJ00xMycpLCAoVCAkZmluICdOMTMnKSkKICAkdmFwQ29sZEluID0gRmlyc3QtTnVtYmVyVGV4dCBAKChUICRvdXQgJ1ExNScpLCAoVCAkb3V0ICdSMTUnKSwgKFQgJGZpbiAnTzEzJyksIChUICRmaW4gJ1AxMycpKQogICR2YXBDb2xkT3V0ID0gRmlyc3QtTnVtYmVyVGV4dCBAKChUICRvdXQgJ1MxNScpLCAoVCAkb3V0ICdUMTUnKSwgKFQgJGZpbiAnUTEzJyksIChUICRmaW4gJ1IxMycpKQogIFB1dCAkZHMgKEFkZHIgJGlzR3BoZSAnQzE3JyAnQzE2JykgJHZhcEhvdEluCiAgUHV0ICRkcyAoQWRkciAkaXNHcGhlICdFMTcnICdFMTYnKSAkdmFwSG90T3V0CiAgUHV0ICRkcyAoQWRkciAkaXNHcGhlICdHMTcnICdHMTYnKSAkdmFwQ29sZEluCiAgUHV0ICRkcyAoQWRkciAkaXNHcGhlICdJMTcnICdJMTYnKSAkdmFwQ29sZE91dAoKICBpZiAoJGlzR3BoZSkgewogICAgUHV0ICRkcyAnQzE2JyAoRmlyc3QtTnVtYmVyVGV4dCBAKChUICRvdXQgJ0cxOCcpLCAoVCAkb3V0ICdIMTgnKSwgKFQgJGZpbiAnSTE4JyksIChUICRmaW4gJ0oxOCcpLCAoVCAkYXBpICdNMzQnKSkpCiAgICBQdXQgJGRzICdHMTYnIChGaXJzdC1OdW1iZXJUZXh0IEAoKFQgJG91dCAnUTE4JyksIChUICRvdXQgJ1IxOCcpLCAoVCAkZmluICdPMTgnKSwgKFQgJGZpbiAnUDE4JyksIChUICRhcGkgJ1czNCcpKSkKICAgIFB1dCAkZHMgJ0oxNicgKCcoJyArIChUICRhcGkgJ0kzNCcpICsgJyknKQogICAgUHV0ICRkcyAnQzE5JyAoRmlyc3QtTnVtYmVyVGV4dCBAKChUICRvdXQgJ0cxNycpLCAoVCAkb3V0ICdIMTcnKSwgKFQgJGZpbiAnSTE3JyksIChUICRmaW4gJ0oxNycpLCAoVCAkYXBpICdNMzMnKSkpCiAgICBQdXQgJGRzICdHMTknIChGaXJzdC1OdW1iZXJUZXh0IEAoKFQgJG91dCAnUTE3JyksIChUICRvdXQgJ1IxNycpLCAoVCAkZmluICdPMTcnKSwgKFQgJGZpbiAnUDE3JyksIChUICRhcGkgJ1czMycpKSkKICAgIFB1dCAkZHMgJ0oxOScgKCcoJyArIChUICRhcGkgJ0kzMycpICsgJyknKQogIH0gZWxzZSB7CiAgICBQdXQgJGRzICdDMTcnIChGaXJzdC1OdW1iZXJUZXh0IEAoKFQgJG91dCAnRzE4JyksIChUICRvdXQgJ0gxOCcpLCAoVCAkZmluICdJMTgnKSwgKFQgJGZpbiAnSjE4JyksIChUICRhcGkgJ1IzNCcpKSkKICAgIFB1dCAkZHMgJ0cxNycgKEZpcnN0LU51bWJlclRleHQgQCgoVCAkb3V0ICdRMTgnKSwgKFQgJG91dCAnUjE4JyksIChUICRmaW4gJ08xOCcpLCAoVCAkZmluICdQMTgnKSwgKFQgJGFwaSAnQUIzNCcpKSkKICAgIFB1dCAkZHMgJ0oxNycgKCcoJyArIChUICRhcGkgJ0kzNCcpICsgJyknKQogICAgUHV0ICRkcyAnQzE4JyAoRmlyc3QtTnVtYmVyVGV4dCBAKChUICRvdXQgJ0cxNycpLCAoVCAkb3V0ICdIMTcnKSwgKFQgJGZpbiAnSTE3JyksIChUICRmaW4gJ0oxNycpLCAoVCAkYXBpICdNMzMnKSkpCiAgICBQdXQgJGRzICdHMTgnIChGaXJzdC1OdW1iZXJUZXh0IEAoKFQgJG91dCAnUTE3JyksIChUICRvdXQgJ1IxNycpLCAoVCAkZmluICdPMTcnKSwgKFQgJGZpbiAnUDE3JyksIChUICRhcGkgJ1czMycpKSkKICAgIFB1dCAkZHMgJ0oxOCcgKCcoJyArIChUICRhcGkgJ0kzMycpICsgJyknKQogIH0KCiAgJHJlZkhvdCA9IEZpcnN0LU51bWJlclRleHQgQCgoVCAkZmluICdJMTUnKSwgKFQgJGZpbiAnSjE1JyksIChUICRvdXQgJ0cxNicpLCAoVCAkb3V0ICdIMTYnKSwgKFQgJGFwaSAnTTIyJykpCiAgJHJlZkNvbGQgPSBGaXJzdC1OdW1iZXJUZXh0IEAoKFQgJGZpbiAnTzE1JyksIChUICRmaW4gJ1AxNScpLCAoVCAkb3V0ICdRMTYnKSwgKFQgJG91dCAnUjE2JyksIChUICRhcGkgJ1cyMicpKQogIFB1dCAkZHMgKEFkZHIgJGlzR3BoZSAnQzIzJyAnQzIyJykgJHJlZkhvdAogIFB1dCAkZHMgKEFkZHIgJGlzR3BoZSAnRzIzJyAnRzIyJykgJHJlZkNvbGQKICBQdXQgJGRzIChBZGRyICRpc0dwaGUgJ0oyMycgJ0oyMicpIChUICRhcGkgJ0kyMicpCgogIFB1dERhc2hOdW1iZXIgJGRzIChBZGRyICRpc0dwaGUgJ0MyNCcgJ0MyMycpIChGaXJzdC1OdW1iZXJUZXh0IEAoKFQgJGFwaSAnTDI1JyksIChUICRhcGkgJ1EyNScpLCAoVCAkYXBpICdNMjUnKSkpCiAgUHV0RGFzaE51bWJlciAkZHMgKEFkZHIgJGlzR3BoZSAnRTI0JyAnRTIzJykgKEZpcnN0LU51bWJlclRleHQgQCgoVCAkYXBpICdNMjUnKSwgKFQgJGFwaSAnUjI1JyksIChUICRhcGkgJ04yNScpKSkKICBQdXREYXNoTnVtYmVyICRkcyAoQWRkciAkaXNHcGhlICdHMjQnICdHMjMnKSAoRmlyc3QtTnVtYmVyVGV4dCBAKChUICRhcGkgJ1YyNScpLCAoVCAkYXBpICdBQTI1JyksIChUICRhcGkgJ1cyNScpKSkKICBQdXREYXNoTnVtYmVyICRkcyAoQWRkciAkaXNHcGhlICdJMjQnICdJMjMnKSAoRmlyc3QtTnVtYmVyVGV4dCBAKChUICRhcGkgJ1cyNScpLCAoVCAkYXBpICdBQjI1JyksIChUICRhcGkgJ1gyNScpKSkKICBQdXQgJGRzIChBZGRyICRpc0dwaGUgJ0oyNCcgJ0oyMycpIChUICRhcGkgJ0kyNScpCiAgUHV0RGFzaE51bWJlciAkZHMgKEFkZHIgJGlzR3BoZSAnQzI1JyAnQzI0JykgKEZpcnN0LU51bWJlclRleHQgQCgoVCAkYXBpICdMMjYnKSwgKFQgJGFwaSAnUTI2JyksIChUICRhcGkgJ00yNicpKSkKICBQdXREYXNoTnVtYmVyICRkcyAoQWRkciAkaXNHcGhlICdFMjUnICdFMjQnKSAoRmlyc3QtTnVtYmVyVGV4dCBAKChUICRhcGkgJ00yNicpLCAoVCAkYXBpICdSMjYnKSwgKFQgJGFwaSAnTjI2JykpKQogIFB1dERhc2hOdW1iZXIgJGRzIChBZGRyICRpc0dwaGUgJ0cyNScgJ0cyNCcpIChGaXJzdC1OdW1iZXJUZXh0IEAoKFQgJGFwaSAnVjI2JyksIChUICRhcGkgJ0FBMjYnKSwgKFQgJGFwaSAnVzI2JykpKQogIFB1dERhc2hOdW1iZXIgJGRzIChBZGRyICRpc0dwaGUgJ0kyNScgJ0kyNCcpIChGaXJzdC1OdW1iZXJUZXh0IEAoKFQgJGFwaSAnVzI2JyksIChUICRhcGkgJ0FCMjYnKSwgKFQgJGFwaSAnWDI2JykpKQogIFB1dCAkZHMgKEFkZHIgJGlzR3BoZSAnSjI1JyAnSjI0JykgKFQgJGFwaSAnSTI2JykKICBQdXREYXNoTnVtYmVyICRkcyAoQWRkciAkaXNHcGhlICdDMjYnICdDMjUnKSAoRmlyc3QtTnVtYmVyVGV4dCBAKChUICRhcGkgJ0wyNycpLCAoVCAkYXBpICdRMjcnKSwgKFQgJGFwaSAnTTI3JykpKQogIFB1dERhc2hOdW1iZXIgJGRzIChBZGRyICRpc0dwaGUgJ0UyNicgJ0UyNScpIChGaXJzdC1OdW1iZXJUZXh0IEAoKFQgJGFwaSAnTTI3JyksIChUICRhcGkgJ1IyNycpLCAoVCAkYXBpICdOMjcnKSkpCiAgUHV0RGFzaE51bWJlciAkZHMgKEFkZHIgJGlzR3BoZSAnRzI2JyAnRzI1JykgKEZpcnN0LU51bWJlclRleHQgQCgoVCAkYXBpICdWMjcnKSwgKFQgJGFwaSAnQUEyNycpLCAoVCAkYXBpICdXMjcnKSkpCiAgUHV0RGFzaE51bWJlciAkZHMgKEFkZHIgJGlzR3BoZSAnSTI2JyAnSTI1JykgKEZpcnN0LU51bWJlclRleHQgQCgoVCAkYXBpICdXMjcnKSwgKFQgJGFwaSAnQUIyNycpLCAoVCAkYXBpICdYMjcnKSkpCiAgUHV0ICRkcyAoQWRkciAkaXNHcGhlICdKMjYnICdKMjUnKSAoVCAkYXBpICdJMjcnKQogIFB1dERhc2hOdW1iZXIgJGRzIChBZGRyICRpc0dwaGUgJ0MyNycgJ0MyNicpIChGaXJzdC1OdW1iZXJUZXh0IEAoKFQgJGFwaSAnTDI4JyksIChUICRhcGkgJ1EyOCcpLCAoVCAkYXBpICdNMjgnKSkpCiAgUHV0RGFzaE51bWJlciAkZHMgKEFkZHIgJGlzR3BoZSAnRTI3JyAnRTI2JykgKEZpcnN0LU51bWJlclRleHQgQCgoVCAkYXBpICdNMjgnKSwgKFQgJGFwaSAnUjI4JyksIChUICRhcGkgJ04yOCcpKSkKICBQdXREYXNoTnVtYmVyICRkcyAoQWRkciAkaXNHcGhlICdHMjcnICdHMjYnKSAoRmlyc3QtTnVtYmVyVGV4dCBAKChUICRhcGkgJ1YyOCcpLCAoVCAkYXBpICdBQTI4JyksIChUICRhcGkgJ1cyOCcpKSkKICBQdXREYXNoTnVtYmVyICRkcyAoQWRkciAkaXNHcGhlICdJMjcnICdJMjYnKSAoRmlyc3QtTnVtYmVyVGV4dCBAKChUICRhcGkgJ1cyOCcpLCAoVCAkYXBpICdBQjI4JyksIChUICRhcGkgJ1gyOCcpKSkKICBQdXQgJGRzIChBZGRyICRpc0dwaGUgJ0oyNycgJ0oyNicpICdtTi1zL20yJwoKICBQdXQgJGRzIChBZGRyICRpc0dwaGUgJ0MzMycgJ0MzMCcpIChGaXJzdC1OdW1iZXJUZXh0IEAoKFQgJGFwaSAnRjExJyksIChUICRvdXQgJ1MyNycpKSkKICBQdXQgJGRzIChBZGRyICRpc0dwaGUgJ0czMycgJ0czMCcpIChGaXJzdC1OdW1iZXJUZXh0IEAoKFQgJGFwaSAnRjExJyksIChUICRvdXQgJ1MyNycpKSkKICBQdXQgJGRzIChBZGRyICRpc0dwaGUgJ0MzNCcgJ0MzMScpICRwbGF0ZXMKICBQdXQgJGRzIChBZGRyICRpc0dwaGUgJ0czNCcgJ0czMScpICRwbGF0ZXMKICBQdXQgJGRzIChBZGRyICRpc0dwaGUgJ0MzNicgJ0MzMicpIChGaXJzdC1OdW1iZXJUZXh0IEAoKFQgJGZpbiAnSTMxJyksIChUICRmaW4gJ0ozMScpLCAoVCAkb3V0ICdHMjgnKSwgKFQgJG91dCAnSDI4JyksIChUICRhcGkgJ000MScpKSkKICBQdXQgJGRzIChBZGRyICRpc0dwaGUgJ0MzNycgJ0MzMycpIChGaXJzdC1OdW1iZXJUZXh0IEAoKFQgJG91dCAnUzI0JyksIChUICRvdXQgJ1QyNCcpLCAoVCAkYXBpICdNNDInKSkpCiAgUHV0ICRkcyAoQWRkciAkaXNHcGhlICdHMzcnICdHMzMnKSAoRmlyc3QtTnVtYmVyVGV4dCBAKChUICRvdXQgJ1MyNScpLCAoVCAkb3V0ICdUMjUnKSwgKFQgJGFwaSAnVjQyJyksIChUICRmaW4gJ0kyOScpLCAoVCAkZmluICdKMjknKSkpCiAgUHV0ICRkcyAoQWRkciAkaXNHcGhlICdDMzgnICdDMzQnKSAoRmlyc3QtTnVtYmVyVGV4dCBAKChUICRvdXQgJ1MyOCcpLCAoVCAkb3V0ICdUMjgnKSwgKFQgJGFwaSAnTTQzJykpKQogIFB1dCAkZHMgKEFkZHIgJGlzR3BoZSAnQzM5JyAnQzM1JykgKEZpcnN0LU5vbkJsYW5rIEAoKFQgJGFwaSAnTDQ0JyksICdBSVNJIDMxNicpKQoKICAkcGFyYWxsZWwgPSBGaXJzdC1Vc2VmdWwgQCgoVCAkYXBpICdOMTAnKSwgJzEnKQogICRzZXJpZXMgPSBGaXJzdC1Vc2VmdWwgQCgoVCAkYXBpICdRMTAnKSwgJzEnKQogICRwTnVtID0gTnVtICRwYXJhbGxlbAogICRzTnVtID0gTnVtICRzZXJpZXMKICAkdG90YWxVbml0cyA9IGlmICgkbnVsbCAtbmUgJHBOdW0gLWFuZCAkbnVsbCAtbmUgJHNOdW0pIHsgW2ludF0oJHBOdW0gKiAkc051bSkgfSBlbHNlIHsgJzEnIH0KCiAgJGhvdFByZXNzdXJlID0gRmlyc3QtTnVtYmVyVGV4dCBAKChUICRvdXQgJ0cxNycpLCAoVCAkb3V0ICdIMTcnKSwgKFQgJGZpbiAnSTE3JyksIChUICRmaW4gJ0oxNycpLCAoVCAkYXBpICdNMzMnKSkKICAkY29sZFByZXNzdXJlID0gRmlyc3QtTnVtYmVyVGV4dCBAKChUICRvdXQgJ1ExNycpLCAoVCAkb3V0ICdSMTcnKSwgKFQgJGZpbiAnTzE3JyksIChUICRmaW4gJ1AxNycpLCAoVCAkYXBpICdXMzMnKSkKICAkcHJlc3N1cmVCYXNlID0gTWF4LU51bWJlciBAKCRob3RQcmVzc3VyZSwgJGNvbGRQcmVzc3VyZSkgMTAuMAogIGlmICgkcHJlc3N1cmVCYXNlIC1sZSAxMC4wKSB7CiAgICAkZGVzaWduUHJlc3N1cmUgPSAxMAogIH0gZWxzZSB7CiAgICAkZGVzaWduUHJlc3N1cmUgPSBbTWF0aF06OkNlaWxpbmcoW01hdGhdOjpNYXgoJHByZXNzdXJlQmFzZSArIDIuMCwgJHByZXNzdXJlQmFzZSAqIDEuMSkpCiAgfQogICR0ZXN0UHJlc3N1cmUgPSBbTWF0aF06OlJvdW5kKCRkZXNpZ25QcmVzc3VyZSAqIDEuMywgMSkKICAkbWluVGVtcCA9IFtNYXRoXTo6Rmxvb3IoKE1pbi1OdW1iZXIgQCgoVCAkb3V0ICdHMTQnKSwgKFQgJG91dCAnSDE0JyksIChUICRvdXQgJ00xNCcpLCAoVCAkb3V0ICdOMTQnKSwgKFQgJG91dCAnUTE0JyksIChUICRvdXQgJ1IxNCcpLCAoVCAkb3V0ICdTMTQnKSwgKFQgJG91dCAnVDE0JyksIChUICRmaW4gJ0kxNCcpLCAoVCAkZmluICdKMTQnKSwgKFQgJGZpbiAnTTE0JyksIChUICRmaW4gJ04xNCcpLCAoVCAkZmluICdPMTQnKSwgKFQgJGZpbiAnUDE0JyksIChUICRmaW4gJ1ExNCcpLCAoVCAkZmluICdSMTQnKSwgKFQgJGFwaSAnTTIyJyksIChUICRhcGkgJ1IyMicpLCAoVCAkYXBpICdXMjInKSwgKFQgJGFwaSAnQUIyMicpKSAwLjApKQogICRtYXhUZW1wID0gW01hdGhdOjpDZWlsaW5nKChNYXgtTnVtYmVyIEAoKFQgJG91dCAnRzE0JyksIChUICRvdXQgJ0gxNCcpLCAoVCAkb3V0ICdNMTQnKSwgKFQgJG91dCAnTjE0JyksIChUICRvdXQgJ1ExNCcpLCAoVCAkb3V0ICdSMTQnKSwgKFQgJG91dCAnUzE0JyksIChUICRvdXQgJ1QxNCcpLCAoVCAkZmluICdJMTQnKSwgKFQgJGZpbiAnSjE0JyksIChUICRmaW4gJ00xNCcpLCAoVCAkZmluICdOMTQnKSwgKFQgJGZpbiAnTzE0JyksIChUICRmaW4gJ1AxNCcpLCAoVCAkZmluICdRMTQnKSwgKFQgJGZpbiAnUjE0JyksIChUICRhcGkgJ00yMicpLCAoVCAkYXBpICdSMjInKSwgKFQgJGFwaSAnVzIyJyksIChUICRhcGkgJ0FCMjInKSkgMTAwLjApKQoKICBpZiAoJGlzR3BoZSkgewogICAgUHV0ICRkcyAnQzE4JyAoRmlyc3QtTnVtYmVyVGV4dCBAKChUICRmaW4gJ0oyMCcpLCAoVCAkb3V0ICdIMjAnKSkpCiAgICBQdXQgJGRzICdFMTgnIChGaXJzdC1OdW1iZXJUZXh0IEAoKFQgJGZpbiAnTjIwJyksIChUICRvdXQgJ04yMCcpKSkKICAgIFB1dCAkZHMgJ0cxOCcgKEZpcnN0LU51bWJlclRleHQgQCgoVCAkZmluICdQMjAnKSwgKFQgJG91dCAnUjIwJykpKQogICAgUHV0ICRkcyAnSTE4JyAoRmlyc3QtTnVtYmVyVGV4dCBAKChUICRmaW4gJ1IyMCcpLCAoVCAkb3V0ICdUMjAnKSkpCiAgICBQdXQgJGRzICdKMTgnIChVbml0LVRleHQgKEZpcnN0LU5vbkJsYW5rIEAoKFQgJGZpbiAnSTIwJyksICdtL3MnKSkpCiAgICBQdXQgJGRzICdDMzEnICRwYXJhbGxlbAogICAgUHV0ICRkcyAnRjMxJyAkc2VyaWVzCiAgICBQdXQgJGRzICdHMzEnICR0b3RhbFVuaXRzCiAgICBQdXQgJGRzICdDMzInICRtb2RlbAogICAgUHV0ICRkcyAnRjM1JyAoVCAkYXBpICdMMzknKQogICAgUHV0ICRkcyAnQzQxJyAoKFQgJGFwaSAnTjM3JykgKyAnIHggJyArIChUICRhcGkgJ1EzNycpKQogICAgUHV0ICRkcyAnRzQxJyAoKFQgJGFwaSAnWDM3JykgKyAnIHggJyArIChBZGQtT25lIChUICRhcGkgJ0FBMzcnKSkpCiAgICBQdXQgJGRzICdDNDInICgoVCAkYXBpICdJNDknKSArICcgJyArIChUICRhcGkgJ0o0OScpKQogICAgUHV0ICRkcyAnRzQyJyAoKFQgJGFwaSAnUzQ5JykgKyAnICcgKyAoVCAkYXBpICdUNDknKSkKICAgIFB1dCAkZHMgJ0M0MCcgJ05CUicKICAgIFB1dCAkZHMgJ0M0MicgKEdwaGUtQ29ubmVjdGlvblRleHQgJG1vZGVsKQogICAgUHV0ICRkcyAnRzQyJyAoR3BoZS1Db25uZWN0aW9uVGV4dCAkbW9kZWwpCiAgICBQdXQgJGRzICdDNDMnICdTdGFpbmxlc3Mgc3RlZWwnCiAgICBQdXQgJGRzICdENDUnICRtaW5UZW1wCiAgICBQdXQgJGRzICdGNDUnICcvJwogICAgUHV0ICRkcyAnSDQ1JyAkbWF4VGVtcAogICAgUHV0ICRkcyAnSjQ1JyAnKERlZyBDKScKICAgIFB1dCAkZHMgJ0U0NicgJGRlc2lnblByZXNzdXJlCiAgICBQdXQgJGRzICdGNDYnICcvJwogICAgUHV0ICRkcyAnRzQ2JyAkdGVzdFByZXNzdXJlCiAgICBQdXQgJGRzICdKNDYnICcoYmFyRyknCiAgfSBlbHNlIHsKICAgICRzb2xkZXIgPSBEaXNwbGF5LVNvbGRlciAkbWF0ZXJpYWxPdmVycmlkZSAkbW9kZWwKICAgIFB1dCAkZHMgJ0MzNicgJHNvbGRlcgogICAgUHV0ICRkcyAnQzM3JyAoVCAkYXBpICdOMzcnKQogICAgUHV0ICRkcyAnRTM3JyAoVCAkYXBpICdRMzcnKQogICAgUHV0ICRkcyAnRzM3JyAoVCAkYXBpICdYMzcnKQogICAgUHV0ICRkcyAnSTM3JyAoQWRkLU9uZSAoVCAkYXBpICdBQTM3JykpCiAgICBQdXQgJGRzICdDMzgnICRwYXJhbGxlbAogICAgUHV0ICRkcyAnRjM4JyAkc2VyaWVzCiAgICBQdXQgJGRzICdHMzgnICR0b3RhbFVuaXRzCiAgICBQdXQgJGRzICdDMzknICdTdGFpbmxlc3Mgc3RlZWwnCiAgICBQdXQgJGRzICdDNDAnICRtaW5UZW1wCiAgICBQdXQgJGRzICdGNDAnICcvJwogICAgUHV0ICRkcyAnRzQwJyAkbWF4VGVtcAogICAgUHV0ICRkcyAnSjQwJyAnKERlZyBDKScKICAgIFB1dCAkZHMgJ0M0MScgJGRlc2lnblByZXNzdXJlCiAgICBQdXQgJGRzICdGNDEnICcvJwogICAgUHV0ICRkcyAnRzQxJyAkdGVzdFByZXNzdXJlCiAgICBQdXQgJGRzICdKNDEnICcoYmFyRyknCiAgfQogIEFwcGx5LUh0cmlPdmVycmlkZXMgJGRzICRpdGVtICRpc0dwaGUKICBDbGVhci1Lbm93blNsYXNoQ2VsbHMgJGRzICRpc0dwaGUKfQoKJGV4Y2VsID0gTmV3LU9iamVjdCAtQ29tT2JqZWN0IEV4Y2VsLkFwcGxpY2F0aW9uCiRleGNlbC5WaXNpYmxlID0gJGZhbHNlCiRleGNlbC5EaXNwbGF5QWxlcnRzID0gJGZhbHNlCnRyeSB7CiAgdHJ5IHsgJGV4Y2VsLkF1dG9tYXRpb25TZWN1cml0eSA9IDMgfSBjYXRjaCB7fQogICRvdXRwdXRzID0gQCgpCiAgZm9yZWFjaCAoJGl0ZW0gaW4gJG1hbmlmZXN0LmZpbGVzKSB7CiAgICAka2luZCA9IGlmICgkaXRlbS5raW5kIC1lcSAnR1BIRScpIHsgJ0dQSEUnIH0gZWxzZSB7ICdCUEhFJyB9CiAgICAkdGVtcGxhdGUgPSBKb2luLVBhdGggJFRlbXBsYXRlRGlyICgka2luZCArICcueGxzbScpCiAgICBpZiAoIShUZXN0LVBhdGggLUxpdGVyYWxQYXRoICR0ZW1wbGF0ZSkpIHsgdGhyb3cgIlRlbXBsYXRlIG5vdCBmb3VuZDogJHRlbXBsYXRlIiB9CiAgICBpZiAoIShUZXN0LVBhdGggLUxpdGVyYWxQYXRoICRpdGVtLnBhdGgpKSB7IHRocm93ICJIVFJJIGZpbGUgbm90IGZvdW5kOiAkKCRpdGVtLnBhdGgpIiB9CgogICAgJHNvdXJjZUJhc2UgPSBbSU8uUGF0aF06OkdldEZpbGVOYW1lV2l0aG91dEV4dGVuc2lvbigkaXRlbS5uYW1lKQogICAgJG1vZGVsRm9yTmFtZSA9IFNhZmUtUGFydCAoTW9kZWwtRnJvbU5hbWUgJHNvdXJjZUJhc2UpICdIVFJJJwogICAgJG1vZGVsRm9yTmFtZSA9IFNhZmUtUGFydCAoU2V0LU1vZGVsUHJlZml4ICRtb2RlbEZvck5hbWUgKE92ZXJyaWRlLVZhbHVlICRpdGVtICdTb2xkZXJpbmcgTWF0ZXJpYWwnICcnKSkgJG1vZGVsRm9yTmFtZQogICAgJHBsYXRlc0Zvck5hbWUgPSBQbGF0ZXMtRnJvbU5hbWUgJHNvdXJjZUJhc2UKICAgIGlmIChbc3RyaW5nXTo6SXNOdWxsT3JXaGl0ZVNwYWNlKCRwbGF0ZXNGb3JOYW1lKSkgewogICAgICAkZmlsZUJhc2UgPSAkbW9kZWxGb3JOYW1lICsgJ19EYXRhc2hlZXQnCiAgICB9IGVsc2UgewogICAgICAkZmlsZUJhc2UgPSAkbW9kZWxGb3JOYW1lICsgJ18nICsgJHBsYXRlc0Zvck5hbWUgKyAncGxfRGF0YXNoZWV0JwogICAgfQogICAgJG91dFBhdGggPSBVbmlxdWUtUGF0aCAoSm9pbi1QYXRoICRPdXRwdXREaXIgKCRmaWxlQmFzZSArICcueGxzbScpKQogICAgQ29weS1JdGVtIC1MaXRlcmFsUGF0aCAkdGVtcGxhdGUgLURlc3RpbmF0aW9uICRvdXRQYXRoIC1Gb3JjZQoKICAgICR0YXJnZXQgPSAkZXhjZWwuV29ya2Jvb2tzLk9wZW4oJG91dFBhdGgpCiAgICAkc291cmNlID0gJGV4Y2VsLldvcmtib29rcy5PcGVuKCRpdGVtLnBhdGgpCiAgICB0cnkgewogICAgICBBcHBseS1IdHJpVmFsdWVzICR0YXJnZXQgJHNvdXJjZSAkaXRlbSAka2luZAogICAgICAkZXhjZWwuQ2FsY3VsYXRlRnVsbFJlYnVpbGQoKQogICAgICAkdGFyZ2V0LlNhdmVBcygkb3V0UGF0aCwgNTIpCiAgICB9IGZpbmFsbHkgewogICAgICAkc291cmNlLkNsb3NlKCRmYWxzZSkKICAgICAgJHRhcmdldC5DbG9zZSgkdHJ1ZSkKICAgIH0KICAgICRvdXRwdXRzICs9ICRvdXRQYXRoCiAgfQoKICBpZiAoJG91dHB1dHMuQ291bnQgLWVxIDEpIHsKICAgIFtwc2N1c3RvbW9iamVjdF1AeyB0eXBlPSdmaWxlJzsgcGF0aD0kb3V0cHV0c1swXSB9IHwgQ29udmVydFRvLUpzb24gLUNvbXByZXNzCiAgfSBlbHNlIHsKICAgICR6aXBQYXRoID0gSm9pbi1QYXRoICRPdXRwdXREaXIgJ0hUUklfRGF0YXNoZWV0cy56aXAnCiAgICBpZiAoVGVzdC1QYXRoIC1MaXRlcmFsUGF0aCAkemlwUGF0aCkgeyBSZW1vdmUtSXRlbSAtTGl0ZXJhbFBhdGggJHppcFBhdGggLUZvcmNlIH0KICAgIENvbXByZXNzLUFyY2hpdmUgLUxpdGVyYWxQYXRoICRvdXRwdXRzIC1EZXN0aW5hdGlvblBhdGggJHppcFBhdGggLUZvcmNlCiAgICBbcHNjdXN0b21vYmplY3RdQHsgdHlwZT0nemlwJzsgcGF0aD0kemlwUGF0aCB9IHwgQ29udmVydFRvLUpzb24gLUNvbXByZXNzCiAgfQp9IGZpbmFsbHkgewogICRleGNlbC5RdWl0KCkKICBbU3lzdGVtLlJ1bnRpbWUuSW50ZXJvcFNlcnZpY2VzLk1hcnNoYWxdOjpSZWxlYXNlQ29tT2JqZWN0KCRleGNlbCkgfCBPdXQtTnVsbAp9DQo=';

function writeHtriExcelCopyScript(tempRoot) {
  const psPath = path.join(tempRoot, 'htri_excel_copy.ps1');
  let script = Buffer.from(htriExcelCopyPs1Base64, 'base64').toString('utf8');
  const patchScript = (source, pattern, replacement, label) => {
    const patched = source.replace(pattern, replacement);
    if (patched === source) throw new Error(`Failed to patch HTRI ${label} script.`);
    return patched;
  };
  const patchAll = (source, pattern, replacement, label) => {
    if (!pattern.test(source)) throw new Error(`Failed to patch HTRI ${label} script.`);
    pattern.lastIndex = 0;
    return source.replace(pattern, replacement);
  };
  const connectionBlock = "  if ($isGphe) {\r\n    Put $ds 'C40' (Override-Value $item 'Gasket Material' (T $ds 'C40'))\r\n    $plateMaterial = Override-Value $item 'Plate Material' ''\r\n    $frameMaterial = Override-Value $item 'Frame Material' 'C.S'\r\n    if (-not [string]::IsNullOrWhiteSpace([string]$plateMaterial)) { Put $ds 'C39' $plateMaterial }\r\n    Put $ds 'C43' $frameMaterial\r\n    $connectionHot = Override-Value $item 'Connection Hot' ''\r\n    $connectionCold = Override-Value $item 'Connection Cold' ''\r\n    if (-not [string]::IsNullOrWhiteSpace([string]$connectionHot)) { Put $ds 'C42' $connectionHot }\r\n    if (-not [string]::IsNullOrWhiteSpace([string]$connectionCold)) { Put $ds 'G42' $connectionCold }\r\n  }\r\n}";
  script = patchScript(
    script,
    /  if \(\$isGphe\) \{\r?\n    Put \$ds 'C40' \(Override-Value \$item 'Gasket Material' \(T \$ds 'C40'\)\)\r?\n  \}\r?\n\}/,
    connectionBlock,
    'connection override'
  );
  script = patchScript(
    script,
    /Put \$ds 'C43' 'Stainless steel'/,
    "Put $ds 'C43' 'C.S'",
    'GPHE frame material'
  );
  script = patchScript(
    script,
    /\$model = First-NonBlank @\(\(Model-FromName \$base\), \(T \$api 'L10'\), \(T \$fin 'L11'\)\)/,
    "$model = First-NonBlank @((Model-FromName (T $fin 'Y11')), (Model-FromName $base), (Model-FromName (T $api 'L10')), (Model-FromName (T $fin 'L11')), (Model-FromName (T $api 'D6')), (Model-FromName (T $api 'M9')))",
    'model extraction'
  );
  script = patchScript(
    script,
    /function Unit-Text\(\[string\]\$unit\) \{\r?\n  \$u = \(\[string\]\$unit\)\.Trim\(\)\r?\n  if \(\[string\]::IsNullOrWhiteSpace\(\$u\)\) \{ return '' \}\r?\n  if \(\$u\.StartsWith\('\('\) -and \$u\.EndsWith\('\)'\)\) \{ return \$u \}\r?\n  return '\(' \+ \$u \+ '\)'\r?\n\}/,
    "function Unit-Text([string]$unit) {\r\n  $u = ([string]$unit).Trim()\r\n  if ([string]::IsNullOrWhiteSpace($u)) { return '' }\r\n  if ($u.StartsWith('(') -and $u.EndsWith(')')) { return $u }\r\n  return '(' + $u + ')'\r\n}\r\n\r\nfunction Unit-From-Label([string]$text) {\r\n  $s = ([string]$text).Trim()\r\n  if ([string]::IsNullOrWhiteSpace($s)) { return '' }\r\n  if ($s -match ',\\s*([^,()]+)$') { return $matches[1].Trim() }\r\n  if ($s -match '\\(([^()]*)\\)\\s*$') { return $matches[1].Trim() }\r\n  return ''\r\n}",
    'unit label extraction'
  );
  script = patchScript(
    script,
    /function Split-Pair\(\[string\]\$text\) \{\r?\n  \$parts = \(\[string\]\$text\) -split '\\s\*\/\\s\*\|\\s\*,\\s\*'\r?\n  if \(\$parts\.Count -ge 2\) \{ return @\(\$parts\[0\]\.Trim\(\), \$parts\[1\]\.Trim\(\)\) \}\r?\n  return @\(\(\[string\]\$text\)\.Trim\(\), ''\)\r?\n\}/,
    "function Split-Pair([string]$text) {\r\n  $s = ([string]$text).Trim()\r\n  $slash = $s.IndexOf('/')\r\n  if ($slash -ge 0) { return @($s.Substring(0, $slash).Trim(), $s.Substring($slash + 1).Trim()) }\r\n  $comma = $s.IndexOf(',')\r\n  if ($comma -ge 0) { return @($s.Substring(0, $comma).Trim(), $s.Substring($comma + 1).Trim()) }\r\n  return @($s, '')\r\n}",
    'split pair without regex'
  );
  script = patchScript(
    script,
    /function Set-ModelPrefix\(\[string\]\$model, \[string\]\$material\) \{\r?\n  \$m = \(\[string\]\$model\)\.Trim\(\)\r?\n  if \(\$m -notmatch '\^\(MC\|MS\)'\) \{ return \$m \}\r?\n  if \(\$material -match 'Stainless\|SUS\|\^MS\$'\) \{ return \(\$m -replace '\^\(MC\|MS\)', 'MS'\) \}\r?\n  if \(\$material -match 'Copper\|\^MC\$'\) \{ return \(\$m -replace '\^\(MC\|MS\)', 'MC'\) \}\r?\n  return \$m\r?\n\}/,
    "function Set-ModelPrefix([string]$model, [string]$material) {\r\n  $m = ([string]$model).Trim()\r\n  if ($m.Length -lt 2) { return $m }\r\n  $prefix = $m.Substring(0, 2)\r\n  if ($prefix -ne 'MC') {\r\n    if ($prefix -ne 'MS') { return $m }\r\n  }\r\n  if ($m.Length -gt 2) { $rest = $m.Substring(2) } else { $rest = '' }\r\n  $mat = ([string]$material).Trim().ToUpperInvariant()\r\n  if ($mat.IndexOf('STAINLESS') -ge 0) { return 'MS' + $rest }\r\n  if ($mat.IndexOf('SUS') -ge 0) { return 'MS' + $rest }\r\n  if ($mat -eq 'MS') { return 'MS' + $rest }\r\n  if ($mat.IndexOf('COPPER') -ge 0) { return 'MC' + $rest }\r\n  if ($mat -eq 'MC') { return 'MC' + $rest }\r\n  return $m\r\n}",
    'model prefix without regex'
  );
  script = patchScript(
    script,
    /function Display-Solder\(\[string\]\$material, \[string\]\$model\) \{\r?\n  if \(\$material -match 'Stainless\|SUS\|\^MS\$'\) \{ return 'Stainless Steel' \}\r?\n  if \(\$material -match 'Copper\|\^MC\$'\) \{ return 'Copper' \}\r?\n  if \(\$model -match '\^MS'\) \{ return 'Stainless Steel' \}\r?\n  return 'Copper'\r?\n\}/,
    "function Display-Solder([string]$material, [string]$model) {\r\n  $mat = ([string]$material).Trim().ToUpperInvariant()\r\n  if ($mat.IndexOf('STAINLESS') -ge 0) { return 'Stainless Steel' }\r\n  if ($mat.IndexOf('SUS') -ge 0) { return 'Stainless Steel' }\r\n  if ($mat -eq 'MS') { return 'Stainless Steel' }\r\n  if ($mat.IndexOf('COPPER') -ge 0) { return 'Copper' }\r\n  if ($mat -eq 'MC') { return 'Copper' }\r\n  $modelText = ([string]$model).Trim()\r\n  if ($modelText.Length -ge 2) {\r\n    if ($modelText.Substring(0, 2) -eq 'MS') { return 'Stainless Steel' }\r\n  }\r\n  return 'Copper'\r\n}",
    'solder display without regex'
  );
  script = patchAll(script, /\('\(' \+ \(T \$api '([^']+)'\) \+ '\)'\)/g, "(Unit-Text (T $api '$1'))", 'api unit text');
  script = patchScript(
    script,
    /  Put \$ds 'J12' '\(kcal\/hr\)'/,
    "  $heatExchangedUnit = First-NonBlank @((Unit-From-Label (T $api 'A40')), (Unit-From-Label (T $api 'B40')), (Unit-From-Label (T $api 'C40')), (Unit-From-Label (T $api 'D40')), (Unit-From-Label (T $api 'E40')), (Unit-From-Label (T $api 'F40')), (Unit-From-Label (T $api 'G40')), (Unit-From-Label (T $api 'H40')), (Unit-From-Label (T $api 'I40')), (Unit-From-Label (T $api 'J40')), (Unit-From-Label (T $api 'K40')), (T $api 'I40'))\r\n  if ([string]::IsNullOrWhiteSpace([string]$heatExchangedUnit)) { $heatExchangedUnit = 'kcal/hr' }\r\n  Put $ds 'J12' (Unit-Text $heatExchangedUnit)",
    'heat exchanged unit'
  );
  script = patchAll(script, /\[Math\]::Round\(\(\$n \* \$from \/ \$to\), 6\)/g, "[Math]::Round(($n * $from / $to), 3)", 'unit conversion precision');
  script = patchAll(script, /NumberFormat = '0\.######'/g, "NumberFormat = '0.###'", 'unit conversion number format');
  script = patchScript(
    script,
    /    Put \$ds 'D45' \$minTemp\r?\n    Put \$ds 'F45' '\/'\r?\n    Put \$ds 'H45' \$maxTemp\r?\n    Put \$ds 'J45' '\(Deg C\)'\r?\n    Put \$ds 'E46' \$designPressure\r?\n    Put \$ds 'F46' '\/'\r?\n    Put \$ds 'G46' \$testPressure\r?\n    Put \$ds 'J46' '\(barG\)'/,
    "    Put $ds 'D44' $minTemp\r\n    Put $ds 'H44' $maxTemp\r\n    Put $ds 'J44' '(Deg C)'\r\n    Put $ds 'E45' $designPressure\r\n    Put $ds 'G45' $testPressure\r\n    Put $ds 'J45' '(barG)'",
    'GPHE design rows'
  );
  script = patchAll(script, /Addr \$isGphe 'D45' 'C40'/g, "Addr $isGphe 'D44' 'C40'", 'design temperature min address');
  script = patchAll(script, /Addr \$isGphe 'H45' 'G40'/g, "Addr $isGphe 'H44' 'G40'", 'design temperature max address');
  script = patchAll(script, /Addr \$isGphe 'J45' 'J40'/g, "Addr $isGphe 'J44' 'J40'", 'design temperature unit address');
  script = patchAll(script, /Addr \$isGphe 'E46' 'C41'/g, "Addr $isGphe 'E45' 'C41'", 'design pressure min address');
  script = patchAll(script, /Addr \$isGphe 'G46' 'G41'/g, "Addr $isGphe 'G45' 'G41'", 'design pressure max address');
  script = patchAll(script, /Addr \$isGphe 'J46' 'J41'/g, "Addr $isGphe 'J45' 'J41'", 'design pressure unit address');
  fs.writeFileSync(psPath, script, 'utf8');
  return psPath;
}

const mime = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.xls': 'application/vnd.ms-excel',
  '.xlsm': 'application/vnd.ms-excel.sheet.macroEnabled.12',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.zip': 'application/zip'
};

function send(res, code, body, type = 'text/plain; charset=utf-8') {
  res.writeHead(code, { 'Content-Type': type, 'Cache-Control': 'no-store' });
  res.end(body);
}

function safeName(name) {
  return path.basename(String(name || 'upload.xls')).replace(/[\\/:*?"<>|]/g, '_');
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on('data', chunk => chunks.push(chunk));
    req.on('end', () => resolve(Buffer.concat(chunks)));
    req.on('error', reject);
  });
}

function parseMultipart(buffer, contentType) {
  const match = /boundary=(?:(?:"([^"]+)")|([^;]+))/i.exec(contentType || '');
  if (!match) throw new Error('Missing multipart boundary');
  const boundary = Buffer.from('--' + (match[1] || match[2]));
  const parts = [];
  let start = buffer.indexOf(boundary);
  while (start >= 0) {
    start += boundary.length;
    if (buffer[start] === 45 && buffer[start + 1] === 45) break;
    if (buffer[start] === 13 && buffer[start + 1] === 10) start += 2;
    const headerEnd = buffer.indexOf(Buffer.from('\r\n\r\n'), start);
    if (headerEnd < 0) break;
    const headers = buffer.slice(start, headerEnd).toString('utf8');
    let next = buffer.indexOf(boundary, headerEnd + 4);
    if (next < 0) break;
    let dataEnd = next;
    if (buffer[dataEnd - 2] === 13 && buffer[dataEnd - 1] === 10) dataEnd -= 2;
    const disposition = /content-disposition:[^\r\n]*/i.exec(headers);
    if (disposition) {
      const line = disposition[0];
      const nameMatch = /name="([^"]+)"/i.exec(line);
      const fileMatch = /filename="([^"]*)"/i.exec(line);
      const fileStarMatch = /filename\*=([^;]+)/i.exec(line);
      let filename = fileMatch ? fileMatch[1] : '';
      if (!filename && fileStarMatch) {
        filename = fileStarMatch[1].replace(/^UTF-8''/i, '').trim();
        try { filename = decodeURIComponent(filename); } catch {}
      }
      if (nameMatch) parts.push({ name: nameMatch[1], filename, data: buffer.slice(headerEnd + 4, dataEnd), headers });
    }
    start = next;
  }
  return parts;
}

async function handleHtriConvert(req, res) {
  try {
    const tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'datasheet-pilot-'));
    const uploadDir = path.join(tempRoot, 'uploads');
    const outputDir = path.join(tempRoot, 'out');
    fs.mkdirSync(uploadDir, { recursive: true });
    const parts = parseMultipart(await readBody(req), req.headers['content-type']);
    const manifest = { files: [] };
    const overrides = {};
    for (const part of parts) {
      if (!part.filename && part.name && part.name.startsWith('override__')) {
        const indexText = part.name.split('__')[1];
        try { overrides[indexText] = JSON.parse(part.data.toString('utf8') || '{}'); } catch { overrides[indexText] = {}; }
        continue;
      }
      if (!part.filename) continue;
      const [prefix, indexText, kindText] = part.name.split('__');
      if (prefix !== 'htri') continue;
      const filename = safeName(part.filename);
      const filePath = path.join(uploadDir, `${indexText}_${filename}`);
      fs.writeFileSync(filePath, part.data);
      manifest.files.push({ name: filename, path: filePath, kind: kindText === 'GPHE' ? 'GPHE' : 'BPHE', overrides: overrides[indexText] || {} });
    }
    if (!manifest.files.length) throw new Error('No HTRI files were uploaded');
    console.log('[HTRI] convert start', manifest.files.map(file => `${file.kind}:${file.name}`).join(', '));
    const manifestPath = path.join(tempRoot, 'manifest.json');
    fs.writeFileSync(manifestPath, JSON.stringify(manifest), 'utf8');
    const psPath = writeHtriExcelCopyScript(tempRoot);
    const templateDir = fs.existsSync(path.join(root, 'templates')) ? path.join(root, 'templates') : path.join(root, 'template');
    const args = ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', psPath, '-ManifestPath', manifestPath, '-OutputDir', outputDir, '-TemplateDir', templateDir];
    const child = spawn('powershell.exe', args, { windowsHide: true });
    let stdout = '';
    let stderr = '';
    child.stdout.on('data', data => stdout += data.toString());
    child.stderr.on('data', data => stderr += data.toString());
    child.on('close', code => {
      try {
        if (code !== 0) {
          console.error(`[HTRI] PowerShell failed with exit code ${code}`);
          if (stdout.trim()) console.error(`[HTRI] stdout:\n${stdout.trim()}`);
          if (stderr.trim()) console.error(`[HTRI] stderr:\n${stderr.trim()}`);
          throw new Error(stderr || stdout || `PowerShell exited ${code}`);
        }
        const line = stdout.trim().split(/\r?\n/).filter(Boolean).pop();
        const result = JSON.parse(line);
        const outPath = result.path;
        const filename = path.basename(outPath);
        res.writeHead(200, {
          'Content-Type': result.type === 'zip' ? 'application/zip' : 'application/vnd.ms-excel.sheet.macroEnabled.12',
          'Content-Disposition': `attachment; filename="${encodeURIComponent(filename)}"`,
          'Cache-Control': 'no-store'
        });
        fs.createReadStream(outPath).pipe(res);
      } catch (error) {
        const message = String(error.message || error);
        console.error(`[HTRI] conversion failed:\n${message}`);
        send(res, 500, message);
      }
    });
  } catch (error) {
    const message = String(error.message || error);
    console.error(`[HTRI] conversion failed:\n${message}`);
    send(res, 500, message);
  }
}

http.createServer((req, res) => {
  if (req.method === 'POST' && req.url === '/api/htri-convert') {
    handleHtriConvert(req, res);
    return;
  }
  let pathname;
  try {
    pathname = decodeURIComponent(new URL(req.url, `http://localhost:${port}`).pathname);
  } catch (error) {
    send(res, 400, 'Bad request');
    return;
  }
  if (pathname === '/') pathname = '/DataSheet Pilot.html';
  const filePath = path.normalize(path.join(root, pathname));
  if (!filePath.startsWith(root)) {
    send(res, 403, 'Forbidden');
    return;
  }
  fs.readFile(filePath, (error, data) => {
    if (error) {
      send(res, 404, 'not found');
      return;
    }
    res.writeHead(200, {
      'Content-Type': mime[path.extname(filePath).toLowerCase()] || 'application/octet-stream',
      'Cache-Control': 'no-store'
    });
    res.end(data);
  });
}).listen(port, '127.0.0.1', () => {
  console.log(`DataSheet Pilot running at http://127.0.0.1:${port}/`);
});


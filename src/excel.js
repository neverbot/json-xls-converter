import Sheet from './sheet.js';
import { OrderedMap } from 'js-sdsl';
import JSZip from 'jszip';

const templateXLSX =
  'UEsDBAoAAAAAABN7eUK9Z10uNQQAADUEAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbO+7vzw/eG1sIHZlcnNpb249IjEuMCIgZW5jb2Rpbmc9InV0Zi04Ij8+PFR5cGVzIHhtbG5zPSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvcGFja2FnZS8yMDA2L2NvbnRlbnQtdHlwZXMiPjxEZWZhdWx0IEV4dGVuc2lvbj0ieG1sIiBDb250ZW50VHlwZT0iYXBwbGljYXRpb24vdm5kLm9wZW54bWxmb3JtYXRzLW9mZmljZWRvY3VtZW50LnNwcmVhZHNoZWV0bWwuc2hlZXQubWFpbit4bWwiIC8+PERlZmF1bHQgRXh0ZW5zaW9uPSJyZWxzIiBDb250ZW50VHlwZT0iYXBwbGljYXRpb24vdm5kLm9wZW54bWxmb3JtYXRzLXBhY2thZ2UucmVsYXRpb25zaGlwcyt4bWwiIC8+PERlZmF1bHQgRXh0ZW5zaW9uPSJwc21kY3AiIENvbnRlbnRUeXBlPSJhcHBsaWNhdGlvbi92bmQub3BlbnhtbGZvcm1hdHMtcGFja2FnZS5jb3JlLXByb3BlcnRpZXMreG1sIiAvPjxPdmVycmlkZSBQYXJ0TmFtZT0iL2RvY1Byb3BzL2FwcC54bWwiIENvbnRlbnRUeXBlPSJhcHBsaWNhdGlvbi92bmQub3BlbnhtbGZvcm1hdHMtb2ZmaWNlZG9jdW1lbnQuZXh0ZW5kZWQtcHJvcGVydGllcyt4bWwiIC8+PE92ZXJyaWRlIFBhcnROYW1lPSIveGwvc2hhcmVkU3RyaW5ncy54bWwiIENvbnRlbnRUeXBlPSJhcHBsaWNhdGlvbi92bmQub3BlbnhtbGZvcm1hdHMtb2ZmaWNlZG9jdW1lbnQuc3ByZWFkc2hlZXRtbC5zaGFyZWRTdHJpbmdzK3htbCIgLz48T3ZlcnJpZGUgUGFydE5hbWU9Ii94bC9zdHlsZXMueG1sIiBDb250ZW50VHlwZT0iYXBwbGljYXRpb24vdm5kLm9wZW54bWxmb3JtYXRzLW9mZmljZWRvY3VtZW50LnNwcmVhZHNoZWV0bWwuc3R5bGVzK3htbCIgLz48T3ZlcnJpZGUgUGFydE5hbWU9Ii94bC93b3Jrc2hlZXRzL3NoZWV0LnhtbCIgQ29udGVudFR5cGU9ImFwcGxpY2F0aW9uL3ZuZC5vcGVueG1sZm9ybWF0cy1vZmZpY2Vkb2N1bWVudC5zcHJlYWRzaGVldG1sLndvcmtzaGVldCt4bWwiIC8+PE92ZXJyaWRlIFBhcnROYW1lPSIveGwvdGhlbWUvdGhlbWUueG1sIiBDb250ZW50VHlwZT0iYXBwbGljYXRpb24vdm5kLm9wZW54bWxmb3JtYXRzLW9mZmljZWRvY3VtZW50LnRoZW1lK3htbCIgLz48L1R5cGVzPlBLAwQKAAAAAACLdTlIAAAAAAAAAAAAAAAABgAAAF9yZWxzL1BLAwQKAAAAAAATe3lCdJmAA5wCAACcAgAACwAAAF9yZWxzLy5yZWxz77u/PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0idXRmLTgiPz48UmVsYXRpb25zaGlwcyB4bWxucz0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL3BhY2thZ2UvMjAwNi9yZWxhdGlvbnNoaXBzIj48UmVsYXRpb25zaGlwIFR5cGU9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9vZmZpY2VEb2N1bWVudC8yMDA2L3JlbGF0aW9uc2hpcHMvb2ZmaWNlRG9jdW1lbnQiIFRhcmdldD0iL3hsL3dvcmtib29rLnhtbCIgSWQ9IlI0YWEyMmIzMWExYTc0MjkxIiAvPjxSZWxhdGlvbnNoaXAgVHlwZT0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL29mZmljZURvY3VtZW50LzIwMDYvcmVsYXRpb25zaGlwcy9leHRlbmRlZC1wcm9wZXJ0aWVzIiBUYXJnZXQ9Ii9kb2NQcm9wcy9hcHAueG1sIiBJZD0icklkMSIgLz48UmVsYXRpb25zaGlwIFR5cGU9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9wYWNrYWdlLzIwMDYvcmVsYXRpb25zaGlwcy9tZXRhZGF0YS9jb3JlLXByb3BlcnRpZXMiIFRhcmdldD0iL3BhY2thZ2Uvc2VydmljZXMvbWV0YWRhdGEvY29yZS1wcm9wZXJ0aWVzL2VjZmRkMzE0M2YyMTQ4OTA5NWE0NGM3MTExNWI3MjNiLnBzbWRjcCIgSWQ9IlJlZjQ4N2MzZTBjNzQ0YTg3IiAvPjwvUmVsYXRpb25zaGlwcz5QSwMECgAAAAAAi3U5SAAAAAAAAAAAAAAAAAkAAABkb2NQcm9wcy9QSwMECgAAAAAAE3t5Qu9e3149AwAAPQMAABAAAABkb2NQcm9wcy9hcHAueG1s77u/PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0idXRmLTgiPz48YXA6UHJvcGVydGllcyB4bWxuczp2dD0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL29mZmljZURvY3VtZW50LzIwMDYvZG9jUHJvcHNWVHlwZXMiIHhtbG5zOmFwPSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvb2ZmaWNlRG9jdW1lbnQvMjAwNi9leHRlbmRlZC1wcm9wZXJ0aWVzIj48YXA6QXBwbGljYXRpb24+TWljcm9zb2Z0IEV4Y2VsPC9hcDpBcHBsaWNhdGlvbj48YXA6RG9jU2VjdXJpdHk+MDwvYXA6RG9jU2VjdXJpdHk+PGFwOlNjYWxlQ3JvcD5mYWxzZTwvYXA6U2NhbGVDcm9wPjxhcDpIZWFkaW5nUGFpcnM+PHZ0OnZlY3RvciBiYXNlVHlwZT0idmFyaWFudCIgc2l6ZT0iNCI+PHZ0OnZhcmlhbnQ+PHZ0Omxwc3RyPldvcmtzaGVldHM8L3Z0Omxwc3RyPjwvdnQ6dmFyaWFudD48dnQ6dmFyaWFudD48dnQ6aTQ+MTwvdnQ6aTQ+PC92dDp2YXJpYW50Pjx2dDp2YXJpYW50Pjx2dDpscHN0cj5OYW1lZCBSYW5nZXM8L3Z0Omxwc3RyPjwvdnQ6dmFyaWFudD48dnQ6dmFyaWFudD48dnQ6aTQ+MjwvdnQ6aTQ+PC92dDp2YXJpYW50PjwvdnQ6dmVjdG9yPjwvYXA6SGVhZGluZ1BhaXJzPjxhcDpUaXRsZXNPZlBhcnRzPjx2dDp2ZWN0b3IgYmFzZVR5cGU9Imxwc3RyIiBzaXplPSIzIj48dnQ6bHBzdHI+U2hlZXQgMTwvdnQ6bHBzdHI+PHZ0Omxwc3RyPlNoZWV0IDEhUHJpbnRfQXJlYTwvdnQ6bHBzdHI+PHZ0Omxwc3RyPlNoZWV0IDEhUHJpbnRfVGl0bGVzPC92dDpscHN0cj48L3Z0OnZlY3Rvcj48L2FwOlRpdGxlc09mUGFydHM+PC9hcDpQcm9wZXJ0aWVzPlBLAwQKAAAAAACLdTlIAAAAAAAAAAAAAAAACAAAAHBhY2thZ2UvUEsDBAoAAAAAAMWFeUIAAAAAAAAAAAAAAAARAAAAcGFja2FnZS9zZXJ2aWNlcy9QSwMECgAAAAAAxYV5QgAAAAAAAAAAAAAAABoAAABwYWNrYWdlL3NlcnZpY2VzL21ldGFkYXRhL1BLAwQKAAAAAADFhXlCAAAAAAAAAAAAAAAAKgAAAHBhY2thZ2Uvc2VydmljZXMvbWV0YWRhdGEvY29yZS1wcm9wZXJ0aWVzL1BLAwQKAAAAAAATe3lCc4c2yNoBAADaAQAAUQAAAHBhY2thZ2Uvc2VydmljZXMvbWV0YWRhdGEvY29yZS1wcm9wZXJ0aWVzL2VjZmRkMzE0M2YyMTQ4OTA5NWE0NGM3MTExNWI3MjNiLnBzbWRjcO+7vzw/eG1sIHZlcnNpb249IjEuMCIgZW5jb2Rpbmc9InV0Zi04Ij8+PGNvcmVQcm9wZXJ0aWVzIHhtbG5zOmRjPSJodHRwOi8vcHVybC5vcmcvZGMvZWxlbWVudHMvMS4xLyIgeG1sbnM6ZGN0ZXJtcz0iaHR0cDovL3B1cmwub3JnL2RjL3Rlcm1zLyIgeG1sbnM6eHNpPSJodHRwOi8vd3d3LnczLm9yZy8yMDAxL1hNTFNjaGVtYS1pbnN0YW5jZSIgeG1sbnM9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9wYWNrYWdlLzIwMDYvbWV0YWRhdGEvY29yZS1wcm9wZXJ0aWVzIj48ZGN0ZXJtczpjcmVhdGVkIHhzaTp0eXBlPSJkY3Rlcm1zOlczQ0RURiI+MjAxMy0wMy0yNVQxOToyNDozOS44NjYxOTY5WjwvZGN0ZXJtczpjcmVhdGVkPjxkY3Rlcm1zOm1vZGlmaWVkIHhzaTp0eXBlPSJkY3Rlcm1zOlczQ0RURiI+MjAxMy0wMy0yNVQxOToyNDozOS44NjYxOTY5WjwvZGN0ZXJtczptb2RpZmllZD48L2NvcmVQcm9wZXJ0aWVzPlBLAwQKAAAAAACLdTlIAAAAAAAAAAAAAAAAAwAAAHhsL1BLAwQKAAAAAADFhXlCAAAAAAAAAAAAAAAACQAAAHhsL19yZWxzL1BLAwQKAAAAAAATe3lCJ0p8MrwCAAC8AgAAGgAAAHhsL19yZWxzL3dvcmtib29rLnhtbC5yZWxz77u/PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0idXRmLTgiPz48UmVsYXRpb25zaGlwcyB4bWxucz0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL3BhY2thZ2UvMjAwNi9yZWxhdGlvbnNoaXBzIj48UmVsYXRpb25zaGlwIFR5cGU9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9vZmZpY2VEb2N1bWVudC8yMDA2L3JlbGF0aW9uc2hpcHMvc2hhcmVkU3RyaW5ncyIgVGFyZ2V0PSIveGwvc2hhcmVkU3RyaW5ncy54bWwiIElkPSJySWQzIiAvPjxSZWxhdGlvbnNoaXAgVHlwZT0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL29mZmljZURvY3VtZW50LzIwMDYvcmVsYXRpb25zaGlwcy9zdHlsZXMiIFRhcmdldD0iL3hsL3N0eWxlcy54bWwiIElkPSJySWQ0IiAvPjxSZWxhdGlvbnNoaXAgVHlwZT0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL29mZmljZURvY3VtZW50LzIwMDYvcmVsYXRpb25zaGlwcy93b3Jrc2hlZXQiIFRhcmdldD0iL3hsL3dvcmtzaGVldHMvc2hlZXQueG1sIiBJZD0icklkMiIgLz48UmVsYXRpb25zaGlwIFR5cGU9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9vZmZpY2VEb2N1bWVudC8yMDA2L3JlbGF0aW9uc2hpcHMvdGhlbWUiIFRhcmdldD0iL3hsL3RoZW1lL3RoZW1lLnhtbCIgSWQ9InJJZDYiIC8+PC9SZWxhdGlvbnNoaXBzPlBLAwQKAAAAAAATe3lCflKRBZAAAACQAAAAFAAAAHhsL3NoYXJlZFN0cmluZ3MueG1s77u/PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0idXRmLTgiPz48eDpzc3QgY291bnQ9IjAiIHVuaXF1ZUNvdW50PSIwIiB4bWxuczp4PSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvc3ByZWFkc2hlZXRtbC8yMDA2L21haW4iIC8+UEsDBAoAAAAAABN7eUIi2lsreggAAHoIAAANAAAAeGwvc3R5bGVzLnhtbO+7vzw/eG1sIHZlcnNpb249IjEuMCIgZW5jb2Rpbmc9InV0Zi04Ij8+PHg6c3R5bGVTaGVldCB4bWxuczp4PSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvc3ByZWFkc2hlZXRtbC8yMDA2L21haW4iPjx4Om51bUZtdHMgY291bnQ9IjEiPjx4Om51bUZtdCBudW1GbXRJZD0iMCIgZm9ybWF0Q29kZT0iIiAvPjwveDpudW1GbXRzPjx4OmZvbnRzIGNvdW50PSIxIj48eDpmb250Pjx4OnZlcnRBbGlnbiB2YWw9ImJhc2VsaW5lIiAvPjx4OnN6IHZhbD0iMTEiIC8+PHg6Y29sb3IgcmdiPSJGRjAwMDAwMCIgLz48eDpuYW1lIHZhbD0iQ2FsaWJyaSIgLz48eDpmYW1pbHkgdmFsPSIyIiAvPjwveDpmb250PjwveDpmb250cz48eDpmaWxscyBjb3VudD0iMiI+PHg6ZmlsbD48eDpwYXR0ZXJuRmlsbCBwYXR0ZXJuVHlwZT0ibm9uZSIgLz48L3g6ZmlsbD48eDpmaWxsPjx4OnBhdHRlcm5GaWxsIHBhdHRlcm5UeXBlPSJncmF5MTI1IiAvPjwveDpmaWxsPjwveDpmaWxscz48eDpib3JkZXJzIGNvdW50PSIxIj48eDpib3JkZXIgZGlhZ29uYWxVcD0iMCIgZGlhZ29uYWxEb3duPSIwIj48eDpsZWZ0IHN0eWxlPSJub25lIj48eDpjb2xvciByZ2I9IkZGMDAwMDAwIiAvPjwveDpsZWZ0Pjx4OnJpZ2h0IHN0eWxlPSJub25lIj48eDpjb2xvciByZ2I9IkZGMDAwMDAwIiAvPjwveDpyaWdodD48eDp0b3Agc3R5bGU9Im5vbmUiPjx4OmNvbG9yIHJnYj0iRkYwMDAwMDAiIC8+PC94OnRvcD48eDpib3R0b20gc3R5bGU9Im5vbmUiPjx4OmNvbG9yIHJnYj0iRkYwMDAwMDAiIC8+PC94OmJvdHRvbT48eDpkaWFnb25hbCBzdHlsZT0ibm9uZSI+PHg6Y29sb3IgcmdiPSJGRjAwMDAwMCIgLz48L3g6ZGlhZ29uYWw+PC94OmJvcmRlcj48L3g6Ym9yZGVycz48eDpjZWxsU3R5bGVYZnMgY291bnQ9IjIiPjx4OnhmIG51bUZtdElkPSIwIiBmb250SWQ9IjAiIGZpbGxJZD0iMCIgYm9yZGVySWQ9IjAiIGFwcGx5TnVtYmVyRm9ybWF0PSIwIiBhcHBseUZpbGw9IjEiIGFwcGx5Qm9yZGVyPSIwIiBhcHBseUFsaWdubWVudD0iMCIgYXBwbHlQcm90ZWN0aW9uPSIxIj48eDpwcm90ZWN0aW9uIGxvY2tlZD0iMSIgaGlkZGVuPSIwIiAvPjwveDp4Zj48eDp4ZiBudW1GbXRJZD0iMTQiIGZvbnRJZD0iMCIgZmlsbElkPSIwIiBib3JkZXJJZD0iMCIgYXBwbHlOdW1iZXJGb3JtYXQ9IjAiIGFwcGx5RmlsbD0iMSIgYXBwbHlCb3JkZXI9IjAiIGFwcGx5QWxpZ25tZW50PSIwIiBhcHBseVByb3RlY3Rpb249IjEiPjx4OnByb3RlY3Rpb24gbG9ja2VkPSIxIiBoaWRkZW49IjAiIC8+PC94OnhmPjwveDpjZWxsU3R5bGVYZnM+PHg6Y2VsbFhmcyBjb3VudD0iMiI+PHg6eGYgbnVtRm10SWQ9IjAiIGZvbnRJZD0iMCIgZmlsbElkPSIwIiBib3JkZXJJZD0iMCIgeGZJZD0iMCIgYXBwbHlOdW1iZXJGb3JtYXQ9IjAiIGFwcGx5RmlsbD0iMSIgYXBwbHlCb3JkZXI9IjAiIGFwcGx5QWxpZ25tZW50PSIwIiBhcHBseVByb3RlY3Rpb249IjEiPjx4OmFsaWdubWVudCBob3Jpem9udGFsPSJnZW5lcmFsIiB2ZXJ0aWNhbD0iYm90dG9tIiB0ZXh0Um90YXRpb249IjAiIHdyYXBUZXh0PSIwIiBpbmRlbnQ9IjAiIHJlbGF0aXZlSW5kZW50PSIwIiBqdXN0aWZ5TGFzdExpbmU9IjAiIHNocmlua1RvRml0PSIwIiByZWFkaW5nT3JkZXI9IjAiIC8+PHg6cHJvdGVjdGlvbiBsb2NrZWQ9IjEiIGhpZGRlbj0iMCIgLz48L3g6eGY+PHg6eGYgbnVtRm10SWQ9IjE0IiBmb250SWQ9IjAiIGZpbGxJZD0iMCIgYm9yZGVySWQ9IjAiIHhmSWQ9IjAiIGFwcGx5TnVtYmVyRm9ybWF0PSIwIiBhcHBseUZpbGw9IjEiIGFwcGx5Qm9yZGVyPSIwIiBhcHBseUFsaWdubWVudD0iMCIgYXBwbHlQcm90ZWN0aW9uPSIxIj48eDphbGlnbm1lbnQgaG9yaXpvbnRhbD0iZ2VuZXJhbCIgdmVydGljYWw9ImJvdHRvbSIgdGV4dFJvdGF0aW9uPSIwIiB3cmFwVGV4dD0iMCIgaW5kZW50PSIwIiByZWxhdGl2ZUluZGVudD0iMCIganVzdGlmeUxhc3RMaW5lPSIwIiBzaHJpbmtUb0ZpdD0iMCIgcmVhZGluZ09yZGVyPSIwIiAvPjx4OnByb3RlY3Rpb24gbG9ja2VkPSIxIiBoaWRkZW49IjAiIC8+PC94OnhmPjwveDpjZWxsWGZzPjx4OmNlbGxTdHlsZXMgY291bnQ9IjEiPjx4OmNlbGxTdHlsZSBuYW1lPSJOb3JtYWwiIHhmSWQ9IjAiIGJ1aWx0aW5JZD0iMCIgLz48L3g6Y2VsbFN0eWxlcz48L3g6c3R5bGVTaGVldD5QSwMECgAAAAAAxYV5QgAAAAAAAAAAAAAAAAkAAAB4bC90aGVtZS9QSwMECgAAAAAAE3t5QnWxkV67GwAAuxsAABIAAAB4bC90aGVtZS90aGVtZS54bWzvu788P3htbCB2ZXJzaW9uPSIxLjAiIGVuY29kaW5nPSJ1dGYtOCI/PjxhOnRoZW1lIHhtbG5zOmE9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9kcmF3aW5nbWwvMjAwNi9tYWluIiBuYW1lPSJPZmZpY2UgVGhlbWUiPjxhOnRoZW1lRWxlbWVudHM+PGE6Y2xyU2NoZW1lIG5hbWU9Ik9mZmljZSI+PGE6ZGsxPjxhOnN5c0NsciB2YWw9IndpbmRvd1RleHQiIGxhc3RDbHI9IjAwMDAwMCIgLz48L2E6ZGsxPjxhOmx0MT48YTpzeXNDbHIgdmFsPSJ3aW5kb3ciIGxhc3RDbHI9IkZGRkZGRiIgLz48L2E6bHQxPjxhOmRrMj48YTpzcmdiQ2xyIHZhbD0iMUY0OTdEIiAvPjwvYTpkazI+PGE6bHQyPjxhOnNyZ2JDbHIgdmFsPSJFRUVDRTEiIC8+PC9hOmx0Mj48YTphY2NlbnQxPjxhOnNyZ2JDbHIgdmFsPSI0RjgxQkQiIC8+PC9hOmFjY2VudDE+PGE6YWNjZW50Mj48YTpzcmdiQ2xyIHZhbD0iQzA1MDREIiAvPjwvYTphY2NlbnQyPjxhOmFjY2VudDM+PGE6c3JnYkNsciB2YWw9IjlCQkI1OSIgLz48L2E6YWNjZW50Mz48YTphY2NlbnQ0PjxhOnNyZ2JDbHIgdmFsPSI4MDY0QTIiIC8+PC9hOmFjY2VudDQ+PGE6YWNjZW50NT48YTpzcmdiQ2xyIHZhbD0iNEJBQ0M2IiAvPjwvYTphY2NlbnQ1PjxhOmFjY2VudDY+PGE6c3JnYkNsciB2YWw9IkY3OTY0NiIgLz48L2E6YWNjZW50Nj48YTpobGluaz48YTpzcmdiQ2xyIHZhbD0iMDAwMEZGIiAvPjwvYTpobGluaz48YTpmb2xIbGluaz48YTpzcmdiQ2xyIHZhbD0iODAwMDgwIiAvPjwvYTpmb2xIbGluaz48L2E6Y2xyU2NoZW1lPjxhOmZvbnRTY2hlbWUgbmFtZT0iT2ZmaWNlIj48YTptYWpvckZvbnQ+PGE6bGF0aW4gdHlwZWZhY2U9IkNhbWJyaWEiIC8+PGE6ZWEgdHlwZWZhY2U9IiIgLz48YTpjcyB0eXBlZmFjZT0iIiAvPjxhOmZvbnQgc2NyaXB0PSJKcGFuIiB0eXBlZmFjZT0i77yt77yzIO+8sOOCtOOCt+ODg+OCryIgLz48YTpmb250IHNjcmlwdD0iSGFuZyIgdHlwZWZhY2U9IuunkeydgCDqs6DrlJUiIC8+PGE6Zm9udCBzY3JpcHQ9IkhhbnMiIHR5cGVmYWNlPSLlrovkvZMiIC8+PGE6Zm9udCBzY3JpcHQ9IkhhbnQiIHR5cGVmYWNlPSLmlrDntLDmmI7pq5QiIC8+PGE6Zm9udCBzY3JpcHQ9IkFyYWIiIHR5cGVmYWNlPSJUaW1lcyBOZXcgUm9tYW4iIC8+PGE6Zm9udCBzY3JpcHQ9IkhlYnIiIHR5cGVmYWNlPSJUaW1lcyBOZXcgUm9tYW4iIC8+PGE6Zm9udCBzY3JpcHQ9IlRoYWkiIHR5cGVmYWNlPSJUYWhvbWEiIC8+PGE6Zm9udCBzY3JpcHQ9IkV0aGkiIHR5cGVmYWNlPSJOeWFsYSIgLz48YTpmb250IHNjcmlwdD0iQmVuZyIgdHlwZWZhY2U9IlZyaW5kYSIgLz48YTpmb250IHNjcmlwdD0iR3VqciIgdHlwZWZhY2U9IlNocnV0aSIgLz48YTpmb250IHNjcmlwdD0iS2htciIgdHlwZWZhY2U9Ik1vb2xCb3JhbiIgLz48YTpmb250IHNjcmlwdD0iS25kYSIgdHlwZWZhY2U9IlR1bmdhIiAvPjxhOmZvbnQgc2NyaXB0PSJHdXJ1IiB0eXBlZmFjZT0iUmFhdmkiIC8+PGE6Zm9udCBzY3JpcHQ9IkNhbnMiIHR5cGVmYWNlPSJFdXBoZW1pYSIgLz48YTpmb250IHNjcmlwdD0iQ2hlciIgdHlwZWZhY2U9IlBsYW50YWdlbmV0IENoZXJva2VlIiAvPjxhOmZvbnQgc2NyaXB0PSJZaWlpIiB0eXBlZmFjZT0iTWljcm9zb2Z0IFlpIEJhaXRpIiAvPjxhOmZvbnQgc2NyaXB0PSJUaWJ0IiB0eXBlZmFjZT0iTWljcm9zb2Z0IEhpbWFsYXlhIiAvPjxhOmZvbnQgc2NyaXB0PSJUaGFhIiB0eXBlZmFjZT0iTVYgQm9saSIgLz48YTpmb250IHNjcmlwdD0iRGV2YSIgdHlwZWZhY2U9Ik1hbmdhbCIgLz48YTpmb250IHNjcmlwdD0iVGVsdSIgdHlwZWZhY2U9IkdhdXRhbWkiIC8+PGE6Zm9udCBzY3JpcHQ9IlRhbWwiIHR5cGVmYWNlPSJMYXRoYSIgLz48YTpmb250IHNjcmlwdD0iU3lyYyIgdHlwZWZhY2U9IkVzdHJhbmdlbG8gRWRlc3NhIiAvPjxhOmZvbnQgc2NyaXB0PSJPcnlhIiB0eXBlZmFjZT0iS2FsaW5nYSIgLz48YTpmb250IHNjcmlwdD0iTWx5bSIgdHlwZWZhY2U9IkthcnRpa2EiIC8+PGE6Zm9udCBzY3JpcHQ9Ikxhb28iIHR5cGVmYWNlPSJEb2tDaGFtcGEiIC8+PGE6Zm9udCBzY3JpcHQ9IlNpbmgiIHR5cGVmYWNlPSJJc2tvb2xhIFBvdGEiIC8+PGE6Zm9udCBzY3JpcHQ9Ik1vbmciIHR5cGVmYWNlPSJNb25nb2xpYW4gQmFpdGkiIC8+PGE6Zm9udCBzY3JpcHQ9IlZpZXQiIHR5cGVmYWNlPSJUaW1lcyBOZXcgUm9tYW4iIC8+PGE6Zm9udCBzY3JpcHQ9IlVpZ2giIHR5cGVmYWNlPSJNaWNyb3NvZnQgVWlnaHVyIiAvPjwvYTptYWpvckZvbnQ+PGE6bWlub3JGb250PjxhOmxhdGluIHR5cGVmYWNlPSJDYWxpYnJpIiAvPjxhOmVhIHR5cGVmYWNlPSIiIC8+PGE6Y3MgdHlwZWZhY2U9IiIgLz48YTpmb250IHNjcmlwdD0iSnBhbiIgdHlwZWZhY2U9Iu+8re+8syDvvLDjgrTjgrfjg4Pjgq8iIC8+PGE6Zm9udCBzY3JpcHQ9IkhhbmciIHR5cGVmYWNlPSLrp5HsnYAg6rOg65SVIiAvPjxhOmZvbnQgc2NyaXB0PSJIYW5zIiB0eXBlZmFjZT0i5a6L5L2TIiAvPjxhOmZvbnQgc2NyaXB0PSJIYW50IiB0eXBlZmFjZT0i5paw57Sw5piO6auUIiAvPjxhOmZvbnQgc2NyaXB0PSJBcmFiIiB0eXBlZmFjZT0iQXJpYWwiIC8+PGE6Zm9udCBzY3JpcHQ9IkhlYnIiIHR5cGVmYWNlPSJBcmlhbCIgLz48YTpmb250IHNjcmlwdD0iVGhhaSIgdHlwZWZhY2U9IlRhaG9tYSIgLz48YTpmb250IHNjcmlwdD0iRXRoaSIgdHlwZWZhY2U9Ik55YWxhIiAvPjxhOmZvbnQgc2NyaXB0PSJCZW5nIiB0eXBlZmFjZT0iVnJpbmRhIiAvPjxhOmZvbnQgc2NyaXB0PSJHdWpyIiB0eXBlZmFjZT0iU2hydXRpIiAvPjxhOmZvbnQgc2NyaXB0PSJLaG1yIiB0eXBlZmFjZT0iRGF1blBlbmgiIC8+PGE6Zm9udCBzY3JpcHQ9IktuZGEiIHR5cGVmYWNlPSJUdW5nYSIgLz48YTpmb250IHNjcmlwdD0iR3VydSIgdHlwZWZhY2U9IlJhYXZpIiAvPjxhOmZvbnQgc2NyaXB0PSJDYW5zIiB0eXBlZmFjZT0iRXVwaGVtaWEiIC8+PGE6Zm9udCBzY3JpcHQ9IkNoZXIiIHR5cGVmYWNlPSJQbGFudGFnZW5ldCBDaGVyb2tlZSIgLz48YTpmb250IHNjcmlwdD0iWWlpaSIgdHlwZWZhY2U9Ik1pY3Jvc29mdCBZaSBCYWl0aSIgLz48YTpmb250IHNjcmlwdD0iVGlidCIgdHlwZWZhY2U9Ik1pY3Jvc29mdCBIaW1hbGF5YSIgLz48YTpmb250IHNjcmlwdD0iVGhhYSIgdHlwZWZhY2U9Ik1WIEJvbGkiIC8+PGE6Zm9udCBzY3JpcHQ9IkRldmEiIHR5cGVmYWNlPSJNYW5nYWwiIC8+PGE6Zm9udCBzY3JpcHQ9IlRlbHUiIHR5cGVmYWNlPSJHYXV0YW1pIiAvPjxhOmZvbnQgc2NyaXB0PSJUYW1sIiB0eXBlZmFjZT0iTGF0aGEiIC8+PGE6Zm9udCBzY3JpcHQ9IlN5cmMiIHR5cGVmYWNlPSJFc3RyYW5nZWxvIEVkZXNzYSIgLz48YTpmb250IHNjcmlwdD0iT3J5YSIgdHlwZWZhY2U9IkthbGluZ2EiIC8+PGE6Zm9udCBzY3JpcHQ9Ik1seW0iIHR5cGVmYWNlPSJLYXJ0aWthIiAvPjxhOmZvbnQgc2NyaXB0PSJMYW9vIiB0eXBlZmFjZT0iRG9rQ2hhbXBhIiAvPjxhOmZvbnQgc2NyaXB0PSJTaW5oIiB0eXBlZmFjZT0iSXNrb29sYSBQb3RhIiAvPjxhOmZvbnQgc2NyaXB0PSJNb25nIiB0eXBlZmFjZT0iTW9uZ29saWFuIEJhaXRpIiAvPjxhOmZvbnQgc2NyaXB0PSJWaWV0IiB0eXBlZmFjZT0iQXJpYWwiIC8+PGE6Zm9udCBzY3JpcHQ9IlVpZ2giIHR5cGVmYWNlPSJNaWNyb3NvZnQgVWlnaHVyIiAvPjwvYTptaW5vckZvbnQ+PC9hOmZvbnRTY2hlbWU+PGE6Zm10U2NoZW1lIG5hbWU9Ik9mZmljZSI+PGE6ZmlsbFN0eWxlTHN0PjxhOnNvbGlkRmlsbD48YTpzY2hlbWVDbHIgdmFsPSJwaENsciIgLz48L2E6c29saWRGaWxsPjxhOmdyYWRGaWxsIHJvdFdpdGhTaGFwZT0iMSI+PGE6Z3NMc3Q+PGE6Z3MgcG9zPSIwIj48YTpzY2hlbWVDbHIgdmFsPSJwaENsciI+PGE6dGludCB2YWw9IjUwMDAwIiAvPjxhOnNhdE1vZCB2YWw9IjMwMDAwMCIgLz48L2E6c2NoZW1lQ2xyPjwvYTpncz48YTpncyBwb3M9IjM1MDAwIj48YTpzY2hlbWVDbHIgdmFsPSJwaENsciI+PGE6dGludCB2YWw9IjM3MDAwIiAvPjxhOnNhdE1vZCB2YWw9IjMwMDAwMCIgLz48L2E6c2NoZW1lQ2xyPjwvYTpncz48YTpncyBwb3M9IjEwMDAwMCI+PGE6c2NoZW1lQ2xyIHZhbD0icGhDbHIiPjxhOnRpbnQgdmFsPSIxNTAwMCIgLz48YTpzYXRNb2QgdmFsPSIzNTAwMDAiIC8+PC9hOnNjaGVtZUNscj48L2E6Z3M+PC9hOmdzTHN0PjxhOmxpbiBhbmc9IjE2MjAwMDAwIiBzY2FsZWQ9IjEiIC8+PC9hOmdyYWRGaWxsPjxhOmdyYWRGaWxsIHJvdFdpdGhTaGFwZT0iMSI+PGE6Z3NMc3Q+PGE6Z3MgcG9zPSIwIj48YTpzY2hlbWVDbHIgdmFsPSJwaENsciI+PGE6c2hhZGUgdmFsPSI1MTAwMCIgLz48YTpzYXRNb2QgdmFsPSIxMzAwMDAiIC8+PC9hOnNjaGVtZUNscj48L2E6Z3M+PGE6Z3MgcG9zPSI4MDAwMCI+PGE6c2NoZW1lQ2xyIHZhbD0icGhDbHIiPjxhOnNoYWRlIHZhbD0iOTMwMDAiIC8+PGE6c2F0TW9kIHZhbD0iMTMwMDAwIiAvPjwvYTpzY2hlbWVDbHI+PC9hOmdzPjxhOmdzIHBvcz0iMTAwMDAwIj48YTpzY2hlbWVDbHIgdmFsPSJwaENsciI+PGE6c2hhZGUgdmFsPSI5NDAwMCIgLz48YTpzYXRNb2QgdmFsPSIxMzUwMDAiIC8+PC9hOnNjaGVtZUNscj48L2E6Z3M+PC9hOmdzTHN0PjxhOmxpbiBhbmc9IjE2MjAwMDAwIiBzY2FsZWQ9IjAiIC8+PC9hOmdyYWRGaWxsPjwvYTpmaWxsU3R5bGVMc3Q+PGE6bG5TdHlsZUxzdD48YTpsbiB3PSI5NTI1IiBjYXA9ImZsYXQiIGNtcGQ9InNuZyIgYWxnbj0iY3RyIj48YTpzb2xpZEZpbGw+PGE6c2NoZW1lQ2xyIHZhbD0icGhDbHIiPjxhOnNoYWRlIHZhbD0iOTUwMDAiIC8+PGE6c2F0TW9kIHZhbD0iMTA1MDAwIiAvPjwvYTpzY2hlbWVDbHI+PC9hOnNvbGlkRmlsbD48YTpwcnN0RGFzaCB2YWw9InNvbGlkIiAvPjwvYTpsbj48YTpsbiB3PSIyNTQwMCIgY2FwPSJmbGF0IiBjbXBkPSJzbmciIGFsZ249ImN0ciI+PGE6c29saWRGaWxsPjxhOnNjaGVtZUNsciB2YWw9InBoQ2xyIiAvPjwvYTpzb2xpZEZpbGw+PGE6cHJzdERhc2ggdmFsPSJzb2xpZCIgLz48L2E6bG4+PGE6bG4gdz0iMzgxMDAiIGNhcD0iZmxhdCIgY21wZD0ic25nIiBhbGduPSJjdHIiPjxhOnNvbGlkRmlsbD48YTpzY2hlbWVDbHIgdmFsPSJwaENsciIgLz48L2E6c29saWRGaWxsPjxhOnByc3REYXNoIHZhbD0ic29saWQiIC8+PC9hOmxuPjwvYTpsblN0eWxlTHN0PjxhOmVmZmVjdFN0eWxlTHN0PjxhOmVmZmVjdFN0eWxlPjxhOmVmZmVjdExzdD48YTpvdXRlclNoZHcgYmx1clJhZD0iNDAwMDAiIGRpc3Q9IjIwMDAwIiBkaXI9IjU0MDAwMDAiIHJvdFdpdGhTaGFwZT0iMCI+PGE6c3JnYkNsciB2YWw9IjAwMDAwMCI+PGE6YWxwaGEgdmFsPSIzODAwMCIgLz48L2E6c3JnYkNscj48L2E6b3V0ZXJTaGR3PjwvYTplZmZlY3RMc3Q+PC9hOmVmZmVjdFN0eWxlPjxhOmVmZmVjdFN0eWxlPjxhOmVmZmVjdExzdD48YTpvdXRlclNoZHcgYmx1clJhZD0iNDAwMDAiIGRpc3Q9IjIzMDAwIiBkaXI9IjU0MDAwMDAiIHJvdFdpdGhTaGFwZT0iMCI+PGE6c3JnYkNsciB2YWw9IjAwMDAwMCI+PGE6YWxwaGEgdmFsPSIzNTAwMCIgLz48L2E6c3JnYkNscj48L2E6b3V0ZXJTaGR3PjwvYTplZmZlY3RMc3Q+PC9hOmVmZmVjdFN0eWxlPjxhOmVmZmVjdFN0eWxlPjxhOmVmZmVjdExzdD48YTpvdXRlclNoZHcgYmx1clJhZD0iNDAwMDAiIGRpc3Q9IjIzMDAwIiBkaXI9IjU0MDAwMDAiIHJvdFdpdGhTaGFwZT0iMCI+PGE6c3JnYkNsciB2YWw9IjAwMDAwMCI+PGE6YWxwaGEgdmFsPSIzNTAwMCIgLz48L2E6c3JnYkNscj48L2E6b3V0ZXJTaGR3PjwvYTplZmZlY3RMc3Q+PGE6c2NlbmUzZD48YTpjYW1lcmEgcHJzdD0ib3J0aG9ncmFwaGljRnJvbnQiPjxhOnJvdCBsYXQ9IjAiIGxvbj0iMCIgcmV2PSIwIiAvPjwvYTpjYW1lcmE+PGE6bGlnaHRSaWcgcmlnPSJ0aHJlZVB0IiBkaXI9InQiPjxhOnJvdCBsYXQ9IjAiIGxvbj0iMCIgcmV2PSIxMjAwMDAwIiAvPjwvYTpsaWdodFJpZz48L2E6c2NlbmUzZD48YTpzcDNkPjxhOmJldmVsVCB3PSI2MzUwMCIgaD0iMjU0MDAiIC8+PC9hOnNwM2Q+PC9hOmVmZmVjdFN0eWxlPjwvYTplZmZlY3RTdHlsZUxzdD48YTpiZ0ZpbGxTdHlsZUxzdD48YTpzb2xpZEZpbGw+PGE6c2NoZW1lQ2xyIHZhbD0icGhDbHIiIC8+PC9hOnNvbGlkRmlsbD48YTpncmFkRmlsbCByb3RXaXRoU2hhcGU9IjEiPjxhOmdzTHN0PjxhOmdzIHBvcz0iMCI+PGE6c2NoZW1lQ2xyIHZhbD0icGhDbHIiPjxhOnRpbnQgdmFsPSI0MDAwMCIgLz48YTpzYXRNb2QgdmFsPSIzNTAwMDAiIC8+PC9hOnNjaGVtZUNscj48L2E6Z3M+PGE6Z3MgcG9zPSI0MDAwMCI+PGE6c2NoZW1lQ2xyIHZhbD0icGhDbHIiPjxhOnRpbnQgdmFsPSI0NTAwMCIgLz48YTpzaGFkZSB2YWw9Ijk5MDAwIiAvPjxhOnNhdE1vZCB2YWw9IjM1MDAwMCIgLz48L2E6c2NoZW1lQ2xyPjwvYTpncz48YTpncyBwb3M9IjEwMDAwMCI+PGE6c2NoZW1lQ2xyIHZhbD0icGhDbHIiPjxhOnNoYWRlIHZhbD0iMjAwMDAiIC8+PGE6c2F0TW9kIHZhbD0iMjU1MDAwIiAvPjwvYTpzY2hlbWVDbHI+PC9hOmdzPjwvYTpnc0xzdD48YTpwYXRoIHBhdGg9ImNpcmNsZSI+PGE6ZmlsbFRvUmVjdCBsPSI1MDAwMCIgdD0iLTgwMDAwIiByPSI1MDAwMCIgYj0iMTgwMDAwIiAvPjwvYTpwYXRoPjwvYTpncmFkRmlsbD48YTpncmFkRmlsbCByb3RXaXRoU2hhcGU9IjEiPjxhOmdzTHN0PjxhOmdzIHBvcz0iMCI+PGE6c2NoZW1lQ2xyIHZhbD0icGhDbHIiPjxhOnRpbnQgdmFsPSI4MDAwMCIgLz48YTpzYXRNb2QgdmFsPSIzMDAwMDAiIC8+PC9hOnNjaGVtZUNscj48L2E6Z3M+PGE6Z3MgcG9zPSIxMDAwMDAiPjxhOnNjaGVtZUNsciB2YWw9InBoQ2xyIj48YTpzaGFkZSB2YWw9IjMwMDAwIiAvPjxhOnNhdE1vZCB2YWw9IjIwMDAwMCIgLz48L2E6c2NoZW1lQ2xyPjwvYTpncz48L2E6Z3NMc3Q+PGE6cGF0aCBwYXRoPSJjaXJjbGUiPjxhOmZpbGxUb1JlY3QgbD0iNTAwMDAiIHQ9IjUwMDAwIiByPSI1MDAwMCIgYj0iNTAwMDAiIC8+PC9hOnBhdGg+PC9hOmdyYWRGaWxsPjwvYTpiZ0ZpbGxTdHlsZUxzdD48L2E6Zm10U2NoZW1lPjwvYTp0aGVtZUVsZW1lbnRzPjxhOm9iamVjdERlZmF1bHRzIC8+PGE6ZXh0cmFDbHJTY2hlbWVMc3QgLz48L2E6dGhlbWU+UEsDBAoAAAAAABN7eUKJ3nBGuwEAALsBAAAPAAAAeGwvd29ya2Jvb2sueG1s77u/PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0idXRmLTgiPz48eDp3b3JrYm9vayB4bWxuczpyPSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvb2ZmaWNlRG9jdW1lbnQvMjAwNi9yZWxhdGlvbnNoaXBzIiB4bWxuczp4PSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvc3ByZWFkc2hlZXRtbC8yMDA2L21haW4iPjx4Ondvcmtib29rUHIgY29kZU5hbWU9IlRoaXNXb3JrYm9vayIgLz48eDpib29rVmlld3M+PHg6d29ya2Jvb2tWaWV3IGZpcnN0U2hlZXQ9IjAiIGFjdGl2ZVRhYj0iMCIgLz48L3g6Ym9va1ZpZXdzPjx4OnNoZWV0cz48eDpzaGVldCBuYW1lPSJTaGVldCAxIiBzaGVldElkPSIyIiByOmlkPSJySWQyIiAvPjwveDpzaGVldHM+PHg6ZGVmaW5lZE5hbWVzIC8+PHg6Y2FsY1ByIGNhbGNJZD0iMTI1NzI1IiAvPjwveDp3b3JrYm9vaz5QSwMECgAAAAAAxYV5QgAAAAAAAAAAAAAAAA4AAAB4bC93b3Jrc2hlZXRzL1BLAQIUAAoAAAAAABN7eUK9Z10uNQQAADUEAAATAAAAAAAAAAAAAAAAAAAAAABbQ29udGVudF9UeXBlc10ueG1sUEsBAhQACgAAAAAAi3U5SAAAAAAAAAAAAAAAAAYAAAAAAAAAAAAQAAAAZgQAAF9yZWxzL1BLAQIUAAoAAAAAABN7eUJ0mYADnAIAAJwCAAALAAAAAAAAAAAAAAAAAIoEAABfcmVscy8ucmVsc1BLAQIUAAoAAAAAAIt1OUgAAAAAAAAAAAAAAAAJAAAAAAAAAAAAEAAAAE8HAABkb2NQcm9wcy9QSwECFAAKAAAAAAATe3lC717fXj0DAAA9AwAAEAAAAAAAAAAAAAAAAAB2BwAAZG9jUHJvcHMvYXBwLnhtbFBLAQIUAAoAAAAAAIt1OUgAAAAAAAAAAAAAAAAIAAAAAAAAAAAAEAAAAOEKAABwYWNrYWdlL1BLAQIUAAoAAAAAAMWFeUIAAAAAAAAAAAAAAAARAAAAAAAAAAAAEAAAAAcLAABwYWNrYWdlL3NlcnZpY2VzL1BLAQIUAAoAAAAAAMWFeUIAAAAAAAAAAAAAAAAaAAAAAAAAAAAAEAAAADYLAABwYWNrYWdlL3NlcnZpY2VzL21ldGFkYXRhL1BLAQIUAAoAAAAAAMWFeUIAAAAAAAAAAAAAAAAqAAAAAAAAAAAAEAAAAG4LAABwYWNrYWdlL3NlcnZpY2VzL21ldGFkYXRhL2NvcmUtcHJvcGVydGllcy9QSwECFAAKAAAAAAATe3lCc4c2yNoBAADaAQAAUQAAAAAAAAAAAAAAAAC2CwAAcGFja2FnZS9zZXJ2aWNlcy9tZXRhZGF0YS9jb3JlLXByb3BlcnRpZXMvZWNmZGQzMTQzZjIxNDg5MDk1YTQ0YzcxMTE1YjcyM2IucHNtZGNwUEsBAhQACgAAAAAAi3U5SAAAAAAAAAAAAAAAAAMAAAAAAAAAAAAQAAAA/w0AAHhsL1BLAQIUAAoAAAAAAMWFeUIAAAAAAAAAAAAAAAAJAAAAAAAAAAAAEAAAACAOAAB4bC9fcmVscy9QSwECFAAKAAAAAAATe3lCJ0p8MrwCAAC8AgAAGgAAAAAAAAAAAAAAAABHDgAAeGwvX3JlbHMvd29ya2Jvb2sueG1sLnJlbHNQSwECFAAKAAAAAAATe3lCflKRBZAAAACQAAAAFAAAAAAAAAAAAAAAAAA7EQAAeGwvc2hhcmVkU3RyaW5ncy54bWxQSwECFAAKAAAAAAATe3lCItpbK3oIAAB6CAAADQAAAAAAAAAAAAAAAAD9EQAAeGwvc3R5bGVzLnhtbFBLAQIUAAoAAAAAAMWFeUIAAAAAAAAAAAAAAAAJAAAAAAAAAAAAEAAAAKIaAAB4bC90aGVtZS9QSwECFAAKAAAAAAATe3lCdbGRXrsbAAC7GwAAEgAAAAAAAAAAAAAAAADJGgAAeGwvdGhlbWUvdGhlbWUueG1sUEsBAhQACgAAAAAAE3t5QonecEa7AQAAuwEAAA8AAAAAAAAAAAAAAAAAtDYAAHhsL3dvcmtib29rLnhtbFBLAQIUAAoAAAAAAMWFeUIAAAAAAAAAAAAAAAAOAAAAAAAAAAAAEAAAAJw4AAB4bC93b3Jrc2hlZXRzL1BLBQYAAAAAEwATANQEAADIOAAAAAA=';
const sheetsFront =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  '<fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="9302"/><workbookPr defaultThemeVersion="124226"/><bookViews><workbookView xWindow="360" yWindow="90" windowWidth="28035" windowHeight="12345" activeTab="0"/></bookViews><sheets>';
const sheetsBack = '</sheets><calcPr calcId="144525"/></workbook>';

const relFront =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';

const relBack =
  '<Relationship Id="rId101" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId100" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId102" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme.xml"/></Relationships>';
const contentTypeFront =
  '<?xml version="1.0" encoding="utf-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />' +
  '<Default Extension="psmdcp" ContentType="application/vnd.openxmlformats-package.core-properties+xml" />' +
  '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml" />' +
  '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" />' +
  '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" />' +
  '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />';
const contentTypeBack =
  '<Override PartName="/xl/theme/theme.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml" /></Types>';
const sharedStringsFront =
  '<?xml version="1.0" encoding="UTF-8"?><x:sst xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="$count">';
const sharedStringsBack = '</x:sst>';
let shareStrings, convertedShareStrings;

async function generateMultiSheets(configs, xlsx) {
  let i = 1;

  for (const config of configs) {
    config.name = config.name ? config.name : 'sheet' + i;
    i++;
    let sheet = new Sheet(config, xlsx, shareStrings, convertedShareStrings);
    await sheet.generate();
    convertedShareStrings = sheet.convertedShareStrings;
  }
}

function generateContentType(configs, xlsx) {
  let workbook = contentTypeFront;

  for (const config of configs) {
    workbook +=
      '<Override PartName="/' +
      config.fileName +
      '" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />';
  }

  workbook += contentTypeBack;
  xlsx.file('[Content_Types].xml', workbook);
}

function generateRel(configs, xlsx) {
  let workbook = relFront;
  let i = 1;

  for (const config of configs) {
    workbook +=
      '<Relationship Id="rId' +
      i +
      '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/' +
      config.name +
      '.xml"/>';
    i++;
  }

  workbook += relBack;
  xlsx.file('xl/_rels/workbook.xml.rels', workbook);
  xlsx.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
  );
}

function generateWorkbook(configs, xlsx) {
  let workbook = sheetsFront;
  let i = 1;

  for (const config of configs) {
    workbook += '<sheet name="' + config.name + '" sheetId="' + i + '" r:id="rId' + i + '"/>';
    i++;
  }

  workbook += sheetsBack;
  xlsx.file('xl/workbook.xml', workbook);
}

function generateSharedStringsFile(xlsx) {
  if (shareStrings.length > 0) {
    let sharedStringsFrontTmp = sharedStringsFront.replace(/\$count/g, shareStrings.length);
    xlsx.file(
      'xl/sharedStrings.xml',
      sharedStringsFrontTmp + convertedShareStrings + sharedStringsBack
    );
  }

  convertedShareStrings = '';
}

// function executeAsync(config, callBack) {
//   return process.nextTick(function () {
//     // var r = exports.execute(config); ??
//     var r = execute(config);
//     callBack(r);
//   });
// }

async function execute(config) {
  let xlsx = await JSZip.loadAsync(templateXLSX, {
    base64: true,
    checkCRC32: false,
  });

  shareStrings = new OrderedMap();
  convertedShareStrings = '';

  let configs = [];

  if (config instanceof Array) {
    configs = config;
  } else {
    configs.push(config);
  }

  await generateMultiSheets(configs, xlsx);
  generateWorkbook(configs, xlsx);
  generateRel(configs, xlsx);
  generateContentType(configs, xlsx);
  generateSharedStringsFile(xlsx);

  const results = await xlsx.generateAsync({
    type: 'uint8array',
    base64: false,
    compression: 'DEFLATE',
  });

  // Parsing error: Deleting local variable in strict mode.
  // delete shareStrings;
  // delete xlsx;
  shareStrings = null;
  xlsx = null;

  return results;
}

export default {
  // executeAsync,
  execute,
};

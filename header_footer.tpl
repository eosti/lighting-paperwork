{% extends "html_table.tpl" %}

{% block before_table %}
{{ generated_page_style | default('') }}
{% endblock before_table %}

{% block before_head_rows %}
<tr class="generatedMarginals">
    <th colspan=42>
        {{ generated_header | default('') }}
    </th>
</tr>
{{ super() }}
{% endblock before_head_rows %}

{% block tbody %}
{{ super() }}
{% block tfoot %}
<tfoot>
    <tr class="generatedMarginals">
        <th colspan=42>
            {{ generated_footer | default('') }}
        </th>
    </tr>
</tfoot>
{% endblock tfoot %}
{% endblock tbody %}

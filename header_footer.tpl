{% extends "html_table.tpl" %}

{% block before_table %}
{% if no_page is sameas false %}
<style>
@page {
  @top-left {
    content: element(topLeft);
  }
  @top-center {
    content: element(topCenter);
  }
  @top-right {
    content: element(topRight);
  }
  @bottom-left {
    {% if pagenum_bottom_left is sameas true %}
    content: "Page " counter(page) " of " counter(pages);
    {{ style_footer_left | default("") }}
    {% else %}
    content: element(bottomLeft);
    {% endif %}
  }
  @bottom-center {
    content: element(bottomCenter);
  }
  @bottom-right {
    {% if pagenum_bottom_right is sameas true %}
    content: "Page " counter(page) " of " counter(pages);
    {{ style_footer_right | default("") }}
    {% else %}
    content: element(bottomRight);
    {% endif %}
  }

  .top-left {
    position: running(topLeft);
  }
  .top-center {
    position: running(topCenter);
  }
  .top-right {
    position: running(topRight);
  }
  .bottom-left {
    position: running(bottomLeft);
  }
  .bottom-center {
    position: running(bottomCenter);
  }
  .bottom-right {
    position: running(bottomRight);
  }
}
</style>
{% endif %}
{% endblock before_table %}

{% block before_head_rows %}
<tr>
    <th colspan=42>
        {{ generated_header | default('') }}
    </th>
</tr>
{{ super() }}
{% endblock before_head_rows %}

{% block tbody %}
{{ super() }}
{% block tfoot %}
<div>
    <tr>
        <th colspan=42>
            {{ generated_footer | default('') }}
        </th>
    </tr>
</div>
{% endblock tfoot %}
{% endblock tbody %}
